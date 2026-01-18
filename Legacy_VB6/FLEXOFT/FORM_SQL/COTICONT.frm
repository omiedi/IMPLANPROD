VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form COTICONT 
   Appearance      =   0  'Flat
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ficha Técnica"
   ClientHeight    =   9375
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   15615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9375
   ScaleWidth      =   15615
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command48 
      Caption         =   "Command48"
      Height          =   255
      Left            =   12720
      TabIndex        =   314
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   7560
      Top             =   4440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command39 
      Caption         =   "Solicitar Cotización"
      Enabled         =   0   'False
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
      Left            =   9480
      TabIndex        =   242
      Top             =   1800
      Width           =   2000
   End
   Begin VB.TextBox TXTNUEVAVERSIONDE 
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
      Height          =   330
      Left            =   13320
      Locked          =   -1  'True
      TabIndex        =   241
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1860
   End
   Begin VB.TextBox TXTLABEL1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
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
      Index           =   2
      Left            =   11640
      Locked          =   -1  'True
      TabIndex        =   240
      TabStop         =   0   'False
      Text            =   "Nueva Versión de"
      Top             =   1800
      Width           =   1665
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2D3C0&
      Height          =   2175
      Left            =   50
      TabIndex        =   25
      Top             =   -50
      Width           =   9255
      Begin VB.TextBox TXTLABEL1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   6720
         Locked          =   -1  'True
         TabIndex        =   320
         TabStop         =   0   'False
         Text            =   "Cód. Espec.Técnica Cliente"
         Top             =   600
         Width           =   2410
      End
      Begin VB.TextBox TXTCODESPECIFTECCLTE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
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
         Left            =   6720
         TabIndex        =   319
         Top             =   1000
         Width           =   2420
      End
      Begin VB.TextBox TEXT1 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   240
         Width           =   900
      End
      Begin VB.CommandButton Command1 
         Caption         =   "."
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
         Left            =   1890
         TabIndex        =   52
         Top             =   240
         Width           =   175
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
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
         Index           =   0
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   51
         TabStop         =   0   'False
         Text            =   "Cantidad de Usos"
         Top             =   1800
         Width           =   1605
      End
      Begin VB.TextBox TXTCANTIUSOS 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
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
         Left            =   4605
         Locked          =   -1  'True
         TabIndex        =   50
         TabStop         =   0   'False
         ToolTipText     =   "Doble Click Para Limpiar Casillero"
         Top             =   1800
         Width           =   1050
      End
      Begin VB.CommandButton Command7 
         Caption         =   "."
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
         Left            =   4395
         TabIndex        =   49
         ToolTipText     =   "6"
         Top             =   1800
         Width           =   175
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
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
         Index           =   41
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   48
         TabStop         =   0   'False
         Text            =   "Producto"
         Top             =   1800
         Width           =   840
      End
      Begin VB.TextBox TXTPRODUCTO 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
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
         Left            =   6600
         TabIndex        =   5
         Top             =   1800
         Width           =   2530
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
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
         Index           =   40
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   47
         TabStop         =   0   'False
         Text            =   "Capacidad en Kgs."
         Top             =   1800
         Width           =   1600
      End
      Begin VB.TextBox TXTCAPACIDAD 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
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
         Left            =   1680
         TabIndex        =   4
         Top             =   1800
         Width           =   1035
      End
      Begin VB.TextBox TXTRAZOCLI 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
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
         Left            =   3000
         TabIndex        =   1
         Top             =   240
         Width           =   6150
      End
      Begin VB.TextBox TXTFECHA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
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
         Left            =   7950
         MaxLength       =   10
         TabIndex        =   38
         Top             =   1410
         Width           =   975
      End
      Begin VB.CommandButton CMDFECHA 
         Caption         =   "."
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
         Left            =   8940
         TabIndex        =   37
         Top             =   1410
         Width           =   175
      End
      Begin VB.CommandButton Command2 
         Caption         =   "."
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
         Left            =   4620
         TabIndex        =   18
         Top             =   1410
         Width           =   175
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
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
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   1410
         Width           =   400
      End
      Begin VB.TextBox TXTLABEL 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
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
         Index           =   15
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         Text            =   "Vendedor"
         Top             =   1410
         Width           =   900
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
         Height          =   330
         Left            =   5235
         Locked          =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         ToolTipText     =   "Doble Click Para Limpiar Casillero"
         Top             =   1410
         Width           =   2655
      End
      Begin VB.TextBox TXTOBSERVA 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
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
         Left            =   1560
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   600
         Width           =   5085
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
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
         Index           =   31
         Left            =   135
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Text            =   "Cliente"
         Top             =   240
         Width           =   840
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
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
         Index           =   11
         Left            =   2175
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Text            =   "R.Social"
         Top             =   240
         Width           =   840
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
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
         Left            =   975
         TabIndex        =   3
         Top             =   1410
         Width           =   900
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
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
         Index           =   12
         Left            =   135
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Text            =   "Cantidad"
         Top             =   1410
         Width           =   840
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
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
         Index           =   13
         Left            =   135
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Text            =   "COTICONT.frx":0000
         Top             =   600
         Width           =   1440
      End
      Begin VB.TextBox TXTLABEL 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
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
         Index           =   14
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Text            =   "Justa"
         Top             =   1410
         Width           =   720
      End
      Begin VB.TextBox TXTJUSTA 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
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
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         ToolTipText     =   "Doble Click Para Aplicar SI"
         Top             =   1410
         Width           =   700
      End
      Begin VB.CommandButton CMD1 
         Caption         =   "."
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
         Left            =   3450
         TabIndex        =   17
         Top             =   1410
         Width           =   175
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2D3C0&
      Height          =   1815
      Left            =   9480
      TabIndex        =   20
      Top             =   -50
      Width           =   6060
      Begin VB.CommandButton Command3 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   2100
         TabIndex        =   324
         Top             =   1080
         Width           =   175
      End
      Begin VB.TextBox TXTLABEL1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
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
         Index           =   1
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   46
         TabStop         =   0   'False
         Text            =   "Revisión"
         Top             =   705
         Width           =   855
      End
      Begin VB.TextBox TXTREVISION 
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
         Height          =   330
         Left            =   5520
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   45
         Top             =   705
         Width           =   405
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
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
         Left            =   5040
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   39
         Top             =   315
         Width           =   885
      End
      Begin VB.TextBox TXTLABEL1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
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
         Index           =   4
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   40
         TabStop         =   0   'False
         Text            =   "N°"
         Top             =   315
         Width           =   375
      End
      Begin VB.TextBox TXTLABEL 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   9
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   44
         TabStop         =   0   'False
         Text            =   "Ficha Técnica"
         Top             =   120
         Width           =   1245
      End
      Begin VB.TextBox TXTDESCRICODCONFECCION 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   2280
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   1080
         Width           =   3645
      End
      Begin VB.TextBox TXTLABEL1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Index           =   3
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Text            =   "COTICONT.frx":001B
         Top             =   1095
         Width           =   1200
      End
      Begin VB.TextBox TXTCODCONFECCION 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   1320
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         ToolTipText     =   "Doble Click Para Limpiar Casillero"
         Top             =   1080
         Width           =   780
      End
      Begin VB.TextBox TXTLABEL1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
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
         Index           =   29
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Text            =   "Código"
         Top             =   200
         Width           =   1065
      End
      Begin VB.TextBox TXTCODIGO 
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
         Height          =   330
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Doble Click Para Borrar"
         Top             =   200
         Width           =   1740
      End
      Begin VB.CommandButton Command44 
         Caption         =   "."
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
         Left            =   2985
         TabIndex        =   22
         Top             =   195
         Visible         =   0   'False
         Width           =   175
      End
      Begin VB.TextBox TXTDESCRICOD 
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
         Height          =   330
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   705
         Width           =   4485
      End
   End
   Begin VB.TextBox TXTCODINT 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   9600
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   1125
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   10560
      OleObjectBlob   =   "COTICONT.frx":0035
      Top             =   8640
   End
   Begin VB.PictureBox Picture6 
      Height          =   9500
      Left            =   120
      ScaleHeight     =   9435
      ScaleWidth      =   16440
      TabIndex        =   41
      Top             =   2160
      Width           =   16500
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00FFFFF0&
         BorderStyle     =   0  'None
         Height          =   15000
         Left            =   -120
         ScaleHeight     =   15000
         ScaleWidth      =   15300
         TabIndex        =   43
         Top             =   120
         Width           =   15300
         Begin TabDlg.SSTab SSTab1 
            Height          =   12600
            Left            =   120
            TabIndex        =   54
            Top             =   -15
            Width           =   15705
            _ExtentX        =   27702
            _ExtentY        =   22225
            _Version        =   393216
            Tabs            =   5
            Tab             =   3
            TabsPerRow      =   5
            TabHeight       =   617
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "Ficha Técnica"
            TabPicture(0)   =   "COTICONT.frx":0269
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "LBLNUEVAFICHATECNICA"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "Label8"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "TXTMICRONAJELINER"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "TXTGRAMAJEFONDO"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "TXTDISTBOCAFDOEIDIT"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "TXTLABEL(33)"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "TXTLABEL(19)"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "TXTLABEL(17)"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).Control(8)=   "TXTLABEL(16)"
            Tab(0).Control(8).Enabled=   0   'False
            Tab(0).Control(9)=   "TXTLABEL(18)"
            Tab(0).Control(9).Enabled=   0   'False
            Tab(0).Control(10)=   "TXTLABEL(5)"
            Tab(0).Control(10).Enabled=   0   'False
            Tab(0).Control(11)=   "TXTLABEL(6)"
            Tab(0).Control(11).Enabled=   0   'False
            Tab(0).Control(12)=   "TXTLABEL(7)"
            Tab(0).Control(12).Enabled=   0   'False
            Tab(0).Control(13)=   "TXTLABEL(3)"
            Tab(0).Control(13).Enabled=   0   'False
            Tab(0).Control(14)=   "TXTLABEL(2)"
            Tab(0).Control(14).Enabled=   0   'False
            Tab(0).Control(15)=   "TXTLABEL(4)"
            Tab(0).Control(15).Enabled=   0   'False
            Tab(0).Control(16)=   "Command33"
            Tab(0).Control(16).Enabled=   0   'False
            Tab(0).Control(17)=   "TXTUNPALLET"
            Tab(0).Control(17).Enabled=   0   'False
            Tab(0).Control(18)=   "TXTLABEL(81)"
            Tab(0).Control(18).Enabled=   0   'False
            Tab(0).Control(19)=   "TXTPALLET"
            Tab(0).Control(19).Enabled=   0   'False
            Tab(0).Control(20)=   "Command32"
            Tab(0).Control(20).Enabled=   0   'False
            Tab(0).Control(21)=   "TXTLABEL(82)"
            Tab(0).Control(21).Enabled=   0   'False
            Tab(0).Control(22)=   "Command20"
            Tab(0).Control(22).Enabled=   0   'False
            Tab(0).Control(23)=   "TXTEMBALAJEFARDO"
            Tab(0).Control(23).Enabled=   0   'False
            Tab(0).Control(24)=   "TXTLABEL(52)"
            Tab(0).Control(24).Enabled=   0   'False
            Tab(0).Control(25)=   "TXTLABEL(39)"
            Tab(0).Control(25).Enabled=   0   'False
            Tab(0).Control(26)=   "Frame3(9)"
            Tab(0).Control(26).Enabled=   0   'False
            Tab(0).Control(27)=   "TXTLABEL(8)"
            Tab(0).Control(27).Enabled=   0   'False
            Tab(0).Control(28)=   "TXTOBSAMPLIA"
            Tab(0).Control(28).Enabled=   0   'False
            Tab(0).Control(29)=   "Command6"
            Tab(0).Control(29).Enabled=   0   'False
            Tab(0).Control(30)=   "Text32"
            Tab(0).Control(30).Enabled=   0   'False
            Tab(0).Control(31)=   "Text31"
            Tab(0).Control(31).Enabled=   0   'False
            Tab(0).Control(32)=   "Text29"
            Tab(0).Control(32).Enabled=   0   'False
            Tab(0).Control(33)=   "TXTLABEL(55)"
            Tab(0).Control(33).Enabled=   0   'False
            Tab(0).Control(34)=   "TXTLABEL(54)"
            Tab(0).Control(34).Enabled=   0   'False
            Tab(0).Control(35)=   "TXTLABEL(53)"
            Tab(0).Control(35).Enabled=   0   'False
            Tab(0).Control(36)=   "TXTCANTICARAS"
            Tab(0).Control(36).Enabled=   0   'False
            Tab(0).Control(37)=   "TXTIMPRESION"
            Tab(0).Control(37).Enabled=   0   'False
            Tab(0).Control(38)=   "Command5"
            Tab(0).Control(38).Enabled=   0   'False
            Tab(0).Control(39)=   "TXTLABEL(42)"
            Tab(0).Control(39).Enabled=   0   'False
            Tab(0).Control(40)=   "TXTLABEL(1)"
            Tab(0).Control(40).Enabled=   0   'False
            Tab(0).Control(41)=   "TXTLABEL(23)"
            Tab(0).Control(41).Enabled=   0   'False
            Tab(0).Control(42)=   "TXTCUERPO_LARGOCM"
            Tab(0).Control(42).Enabled=   0   'False
            Tab(0).Control(43)=   "TXTANCHOCM"
            Tab(0).Control(43).Enabled=   0   'False
            Tab(0).Control(44)=   "Text38"
            Tab(0).Control(44).Enabled=   0   'False
            Tab(0).Control(45)=   "Command22"
            Tab(0).Control(45).Enabled=   0   'False
            Tab(0).Control(46)=   "TXTCOSTURAEXTERNA"
            Tab(0).Control(46).Enabled=   0   'False
            Tab(0).Control(47)=   "TXTCOSTURACONPABILO"
            Tab(0).Control(47).Enabled=   0   'False
            Tab(0).Control(48)=   "TXTLABEL(58)"
            Tab(0).Control(48).Enabled=   0   'False
            Tab(0).Control(49)=   "TXTLABEL(56)"
            Tab(0).Control(49).Enabled=   0   'False
            Tab(0).Control(50)=   "TXTTELAVENTILADA"
            Tab(0).Control(50).Enabled=   0   'False
            Tab(0).Control(51)=   "TXTCUERPOALTOCM"
            Tab(0).Control(51).Enabled=   0   'False
            Tab(0).Control(52)=   "TXTLABEL(43)"
            Tab(0).Control(52).Enabled=   0   'False
            Tab(0).Control(53)=   "Command8"
            Tab(0).Control(53).Enabled=   0   'False
            Tab(0).Control(54)=   "TXTLABEL(57)"
            Tab(0).Control(54).Enabled=   0   'False
            Tab(0).Control(55)=   "Text33"
            Tab(0).Control(55).Enabled=   0   'False
            Tab(0).Control(56)=   "Command21"
            Tab(0).Control(56).Enabled=   0   'False
            Tab(0).Control(57)=   "Command25"
            Tab(0).Control(57).Enabled=   0   'False
            Tab(0).Control(58)=   "Command24"
            Tab(0).Control(58).Enabled=   0   'False
            Tab(0).Control(59)=   "Command23"
            Tab(0).Control(59).Enabled=   0   'False
            Tab(0).Control(60)=   "TXTESLINGAPERIMETRAL"
            Tab(0).Control(60).Enabled=   0   'False
            Tab(0).Control(61)=   "TXTESLINGAANCHO"
            Tab(0).Control(61).Enabled=   0   'False
            Tab(0).Control(62)=   "TXTTIPOCOSTURA"
            Tab(0).Control(62).Enabled=   0   'False
            Tab(0).Control(63)=   "TXTLABEL(61)"
            Tab(0).Control(63).Enabled=   0   'False
            Tab(0).Control(64)=   "TXTLABEL(60)"
            Tab(0).Control(64).Enabled=   0   'False
            Tab(0).Control(65)=   "TXTLABEL(59)"
            Tab(0).Control(65).Enabled=   0   'False
            Tab(0).Control(66)=   "TXT_VPDFESLINGA"
            Tab(0).Control(66).Enabled=   0   'False
            Tab(0).Control(67)=   "TXTESLINGAMEDIDA2CM"
            Tab(0).Control(67).Enabled=   0   'False
            Tab(0).Control(68)=   "TXTESLINGAMEDIDA1CM"
            Tab(0).Control(68).Enabled=   0   'False
            Tab(0).Control(69)=   "TXTESLINGAOJOCM"
            Tab(0).Control(69).Enabled=   0   'False
            Tab(0).Control(70)=   "TXTLABEL(44)"
            Tab(0).Control(70).Enabled=   0   'False
            Tab(0).Control(71)=   "Frame3(12)"
            Tab(0).Control(71).Enabled=   0   'False
            Tab(0).Control(72)=   "Command26"
            Tab(0).Control(72).Enabled=   0   'False
            Tab(0).Control(73)=   "TXTVALVVDDIAMETROCM"
            Tab(0).Control(73).Enabled=   0   'False
            Tab(0).Control(74)=   "TXTLABEL(64)"
            Tab(0).Control(74).Enabled=   0   'False
            Tab(0).Control(75)=   "TXTLABEL(63)"
            Tab(0).Control(75).Enabled=   0   'False
            Tab(0).Control(76)=   "TXT_VPDVDDIAMENTRO"
            Tab(0).Control(76).Enabled=   0   'False
            Tab(0).Control(77)=   "TXTLABEL(62)"
            Tab(0).Control(77).Enabled=   0   'False
            Tab(0).Control(78)=   "TXT_VPDVCLARGOCM"
            Tab(0).Control(78).Enabled=   0   'False
            Tab(0).Control(79)=   "TXT_VPDVCDIAMETROCM"
            Tab(0).Control(79).Enabled=   0   'False
            Tab(0).Control(80)=   "TXTLABEL(45)"
            Tab(0).Control(80).Enabled=   0   'False
            Tab(0).Control(81)=   "TXTVDLARGOCM"
            Tab(0).Control(81).Enabled=   0   'False
            Tab(0).Control(82)=   "TXT_VPDVDLARGO"
            Tab(0).Control(82).Enabled=   0   'False
            Tab(0).Control(83)=   "TXT_VPDVDSLARGOCM"
            Tab(0).Control(83).Enabled=   0   'False
            Tab(0).Control(84)=   "Text19"
            Tab(0).Control(84).Enabled=   0   'False
            Tab(0).Control(85)=   "Command38"
            Tab(0).Control(85).Enabled=   0   'False
            Tab(0).Control(86)=   "TXTLABEL(46)"
            Tab(0).Control(86).Enabled=   0   'False
            Tab(0).Control(87)=   "Command9"
            Tab(0).Control(87).Enabled=   0   'False
            Tab(0).Control(88)=   "Command10"
            Tab(0).Control(88).Enabled=   0   'False
            Tab(0).Control(89)=   "Command11"
            Tab(0).Control(89).Enabled=   0   'False
            Tab(0).Control(90)=   "TXTLABEL(65)"
            Tab(0).Control(90).Enabled=   0   'False
            Tab(0).Control(91)=   "Command27"
            Tab(0).Control(91).Enabled=   0   'False
            Tab(0).Control(92)=   "TXTPOLLERALAMINADA"
            Tab(0).Control(92).Enabled=   0   'False
            Tab(0).Control(93)=   "TXTLABEL(66)"
            Tab(0).Control(93).Enabled=   0   'False
            Tab(0).Control(94)=   "TXTLABEL(67)"
            Tab(0).Control(94).Enabled=   0   'False
            Tab(0).Control(95)=   "TXTLABEL(68)"
            Tab(0).Control(95).Enabled=   0   'False
            Tab(0).Control(96)=   "TXT_VPDPOLLERAGM2"
            Tab(0).Control(96).Enabled=   0   'False
            Tab(0).Control(97)=   "TXT_VPDVDGM2"
            Tab(0).Control(97).Enabled=   0   'False
            Tab(0).Control(98)=   "TXT_VPDVCGM2"
            Tab(0).Control(98).Enabled=   0   'False
            Tab(0).Control(99)=   "TXTLABEL(20)"
            Tab(0).Control(99).Enabled=   0   'False
            Tab(0).Control(100)=   "TXTLABEL(21)"
            Tab(0).Control(100).Enabled=   0   'False
            Tab(0).Control(101)=   "TXTLABEL(22)"
            Tab(0).Control(101).Enabled=   0   'False
            Tab(0).Control(102)=   "TXTTELALAMINADA"
            Tab(0).Control(102).Enabled=   0   'False
            Tab(0).Control(103)=   "TXTLABEL(24)"
            Tab(0).Control(103).Enabled=   0   'False
            Tab(0).Control(104)=   "TXTLABEL(25)"
            Tab(0).Control(104).Enabled=   0   'False
            Tab(0).Control(105)=   "TXTLABEL(47)"
            Tab(0).Control(105).Enabled=   0   'False
            Tab(0).Control(106)=   "Command12"
            Tab(0).Control(106).Enabled=   0   'False
            Tab(0).Control(107)=   "Command13"
            Tab(0).Control(107).Enabled=   0   'False
            Tab(0).Control(108)=   "TXTLABEL(69)"
            Tab(0).Control(108).Enabled=   0   'False
            Tab(0).Control(109)=   "TXTLABEL(70)"
            Tab(0).Control(109).Enabled=   0   'False
            Tab(0).Control(110)=   "Command28"
            Tab(0).Control(110).Enabled=   0   'False
            Tab(0).Control(111)=   "TXT_VPDTELASVALVLAMIN"
            Tab(0).Control(111).Enabled=   0   'False
            Tab(0).Control(112)=   "Frame3(4)"
            Tab(0).Control(112).Enabled=   0   'False
            Tab(0).Control(113)=   "Frame3(8)"
            Tab(0).Control(113).Enabled=   0   'False
            Tab(0).Control(114)=   "TXTLAMCUERPOLAMINADO"
            Tab(0).Control(114).Enabled=   0   'False
            Tab(0).Control(115)=   "TXTLAMFONDOLAMINADO"
            Tab(0).Control(115).Enabled=   0   'False
            Tab(0).Control(116)=   "Command16"
            Tab(0).Control(116).Enabled=   0   'False
            Tab(0).Control(117)=   "TXTLABEL(26)"
            Tab(0).Control(117).Enabled=   0   'False
            Tab(0).Control(118)=   "TXTLABEL(27)"
            Tab(0).Control(118).Enabled=   0   'False
            Tab(0).Control(119)=   "TXTLABEL(28)"
            Tab(0).Control(119).Enabled=   0   'False
            Tab(0).Control(120)=   "TXTLABEL(29)"
            Tab(0).Control(120).Enabled=   0   'False
            Tab(0).Control(121)=   "TXTLABEL(48)"
            Tab(0).Control(121).Enabled=   0   'False
            Tab(0).Control(122)=   "TXTLABEL(49)"
            Tab(0).Control(122).Enabled=   0   'False
            Tab(0).Control(123)=   "TXT_VPDVDESCARGACM"
            Tab(0).Control(123).Enabled=   0   'False
            Tab(0).Control(124)=   "TXT_VPDVDESCSEGCM"
            Tab(0).Control(124).Enabled=   0   'False
            Tab(0).Control(125)=   "TXTPDOCPORTADOCUMENTOTIPO"
            Tab(0).Control(125).Enabled=   0   'False
            Tab(0).Control(126)=   "TXTPDOCPDOCUMENTOUBIC"
            Tab(0).Control(126).Enabled=   0   'False
            Tab(0).Control(127)=   "Command14"
            Tab(0).Control(127).Enabled=   0   'False
            Tab(0).Control(128)=   "Command15"
            Tab(0).Control(128).Enabled=   0   'False
            Tab(0).Control(129)=   "TXTLABEL(71)"
            Tab(0).Control(129).Enabled=   0   'False
            Tab(0).Control(130)=   "TXTLABEL(72)"
            Tab(0).Control(130).Enabled=   0   'False
            Tab(0).Control(131)=   "TXTCINTILLALINER"
            Tab(0).Control(131).Enabled=   0   'False
            Tab(0).Control(132)=   "TXT_VPDPOLLERA"
            Tab(0).Control(132).Enabled=   0   'False
            Tab(0).Control(133)=   "TXTLABEL(73)"
            Tab(0).Control(133).Enabled=   0   'False
            Tab(0).Control(134)=   "TXTLABEL(74)"
            Tab(0).Control(134).Enabled=   0   'False
            Tab(0).Control(135)=   "Command34"
            Tab(0).Control(135).Enabled=   0   'False
            Tab(0).Control(136)=   "Command35"
            Tab(0).Control(136).Enabled=   0   'False
            Tab(0).Control(137)=   "TXTDOCUMENTOUNIDAD"
            Tab(0).Control(137).Enabled=   0   'False
            Tab(0).Control(138)=   "TXTPORTADOCUMEDIDA"
            Tab(0).Control(138).Enabled=   0   'False
            Tab(0).Control(139)=   "Frame3(6)"
            Tab(0).Control(139).Enabled=   0   'False
            Tab(0).Control(140)=   "TXTLABEL(30)"
            Tab(0).Control(140).Enabled=   0   'False
            Tab(0).Control(141)=   "TXTLABEL(34)"
            Tab(0).Control(141).Enabled=   0   'False
            Tab(0).Control(142)=   "TXTLABEL(50)"
            Tab(0).Control(142).Enabled=   0   'False
            Tab(0).Control(143)=   "TXTLABEL(75)"
            Tab(0).Control(143).Enabled=   0   'False
            Tab(0).Control(144)=   "TXTLABEL(76)"
            Tab(0).Control(144).Enabled=   0   'False
            Tab(0).Control(145)=   "TXTLABEL(77)"
            Tab(0).Control(145).Enabled=   0   'False
            Tab(0).Control(146)=   "TXT_VPDANCHOUTILMAMPARO"
            Tab(0).Control(146).Enabled=   0   'False
            Tab(0).Control(147)=   "TXT_VPDDIVISIONLATERALES"
            Tab(0).Control(147).Enabled=   0   'False
            Tab(0).Control(148)=   "TXT_VPDTELAMAMPARO"
            Tab(0).Control(148).Enabled=   0   'False
            Tab(0).Control(149)=   "TXT_VPDDISTANCIABOCAFONDOCM"
            Tab(0).Control(149).Enabled=   0   'False
            Tab(0).Control(150)=   "TXT_VPDANCHO"
            Tab(0).Control(150).Enabled=   0   'False
            Tab(0).Control(151)=   "TXTMICRONAJELINERANTERIOR"
            Tab(0).Control(151).Enabled=   0   'False
            Tab(0).Control(152)=   "Command30"
            Tab(0).Control(152).Enabled=   0   'False
            Tab(0).Control(153)=   "Text56"
            Tab(0).Control(153).Enabled=   0   'False
            Tab(0).Control(154)=   "TXT_VPDALTOLINER"
            Tab(0).Control(154).Enabled=   0   'False
            Tab(0).Control(155)=   "TXTLABEL(80)"
            Tab(0).Control(155).Enabled=   0   'False
            Tab(0).Control(156)=   "Text54"
            Tab(0).Control(156).Enabled=   0   'False
            Tab(0).Control(157)=   "TXTLABEL(79)"
            Tab(0).Control(157).Enabled=   0   'False
            Tab(0).Control(158)=   "TXTLABEL(78)"
            Tab(0).Control(158).Enabled=   0   'False
            Tab(0).Control(159)=   "TXTLINERTIPO"
            Tab(0).Control(159).Enabled=   0   'False
            Tab(0).Control(160)=   "Command18"
            Tab(0).Control(160).Enabled=   0   'False
            Tab(0).Control(161)=   "Command17"
            Tab(0).Control(161).Enabled=   0   'False
            Tab(0).Control(162)=   "TXTLINERLINER"
            Tab(0).Control(162).Enabled=   0   'False
            Tab(0).Control(163)=   "TXT_VPDFUELLELINERCU"
            Tab(0).Control(163).Enabled=   0   'False
            Tab(0).Control(164)=   "TXTLABEL(51)"
            Tab(0).Control(164).Enabled=   0   'False
            Tab(0).Control(165)=   "TXTLABEL(37)"
            Tab(0).Control(165).Enabled=   0   'False
            Tab(0).Control(166)=   "TXTLABEL(36)"
            Tab(0).Control(166).Enabled=   0   'False
            Tab(0).Control(167)=   "TXTLABEL(35)"
            Tab(0).Control(167).Enabled=   0   'False
            Tab(0).Control(168)=   "Frame3(10)"
            Tab(0).Control(168).Enabled=   0   'False
            Tab(0).Control(169)=   "Frame3(7)"
            Tab(0).Control(169).Enabled=   0   'False
            Tab(0).Control(170)=   "TXTDETAETIQUETA2"
            Tab(0).Control(170).Enabled=   0   'False
            Tab(0).Control(171)=   "Command31"
            Tab(0).Control(171).Enabled=   0   'False
            Tab(0).Control(172)=   "Command19"
            Tab(0).Control(172).Enabled=   0   'False
            Tab(0).Control(173)=   "TXTLABEL(83)"
            Tab(0).Control(173).Enabled=   0   'False
            Tab(0).Control(174)=   "TXTDETALLEETIQUETA"
            Tab(0).Control(174).Enabled=   0   'False
            Tab(0).Control(175)=   "TXTLABEL(38)"
            Tab(0).Control(175).Enabled=   0   'False
            Tab(0).Control(176)=   "Frame3(1)"
            Tab(0).Control(176).Enabled=   0   'False
            Tab(0).Control(177)=   "Frame3(13)"
            Tab(0).Control(177).Enabled=   0   'False
            Tab(0).Control(178)=   "Frame3(2)"
            Tab(0).Control(178).Enabled=   0   'False
            Tab(0).Control(179)=   "TXTVCLARGOCM"
            Tab(0).Control(179).Enabled=   0   'False
            Tab(0).Control(180)=   "TXTVALVVCDIAMETROCM"
            Tab(0).Control(180).Enabled=   0   'False
            Tab(0).Control(181)=   "TXTGRAMAJECUERPO"
            Tab(0).Control(181).Enabled=   0   'False
            Tab(0).Control(182)=   "TXTGRAMAJETAPA"
            Tab(0).Control(182).Enabled=   0   'False
            Tab(0).Control(183)=   "TXTPOLLERALARGO"
            Tab(0).Control(183).Enabled=   0   'False
            Tab(0).Control(184)=   "TXTVALVVCLARGOCM"
            Tab(0).Control(184).Enabled=   0   'False
            Tab(0).Control(185)=   "Frame3(11)"
            Tab(0).Control(185).Enabled=   0   'False
            Tab(0).Control(186)=   "TXTCOMPARTCOMPARTIMENTADO"
            Tab(0).Control(186).Enabled=   0   'False
            Tab(0).Control(187)=   "TXT_VPDLARGOMAMAPAROCM"
            Tab(0).Control(187).Enabled=   0   'False
            Tab(0).Control(188)=   "Frame3(16)"
            Tab(0).Control(188).Enabled=   0   'False
            Tab(0).Control(189)=   "Command37"
            Tab(0).Control(189).Enabled=   0   'False
            Tab(0).Control(190)=   "Command36"
            Tab(0).Control(190).Enabled=   0   'False
            Tab(0).Control(191)=   "Frame3(19)"
            Tab(0).Control(191).Enabled=   0   'False
            Tab(0).Control(192)=   "Frame3(20)"
            Tab(0).Control(192).Enabled=   0   'False
            Tab(0).Control(193)=   "Frame3(21)"
            Tab(0).Control(193).Enabled=   0   'False
            Tab(0).Control(194)=   "Frame3(22)"
            Tab(0).Control(194).Enabled=   0   'False
            Tab(0).Control(195)=   "Frame3(23)"
            Tab(0).Control(195).Enabled=   0   'False
            Tab(0).Control(196)=   "Frame3(17)"
            Tab(0).Control(196).Enabled=   0   'False
            Tab(0).Control(197)=   "Command52"
            Tab(0).Control(197).Enabled=   0   'False
            Tab(0).Control(198)=   "Command4"
            Tab(0).Control(198).Enabled=   0   'False
            Tab(0).Control(199)=   "Frame3(24)"
            Tab(0).Control(199).Enabled=   0   'False
            Tab(0).Control(200)=   "Text7"
            Tab(0).Control(200).Enabled=   0   'False
            Tab(0).Control(201)=   "Frame3(3)"
            Tab(0).Control(201).Enabled=   0   'False
            Tab(0).Control(202)=   "Frame3(0)"
            Tab(0).Control(202).Enabled=   0   'False
            Tab(0).Control(203)=   "Text23"
            Tab(0).Control(203).Enabled=   0   'False
            Tab(0).Control(204)=   "Text20"
            Tab(0).Control(204).Enabled=   0   'False
            Tab(0).Control(205)=   "TXT_VPDvcsegcm"
            Tab(0).Control(205).Enabled=   0   'False
            Tab(0).Control(206)=   "TXT_VPDCargaCm"
            Tab(0).Control(206).Enabled=   0   'False
            Tab(0).Control(207)=   "TXT_VPDCargaCm_EDIT"
            Tab(0).Control(207).Enabled=   0   'False
            Tab(0).Control(208)=   "TXT_VPDvcsegcm_EDIT"
            Tab(0).Control(208).Enabled=   0   'False
            Tab(0).Control(209)=   "TXT_VPDVDESCARGACM_EDIT"
            Tab(0).Control(209).Enabled=   0   'False
            Tab(0).Control(210)=   "TXT_VPDVDESCSEGCM_EDIT"
            Tab(0).Control(210).Enabled=   0   'False
            Tab(0).Control(211)=   "Frame3(5)"
            Tab(0).Control(211).Enabled=   0   'False
            Tab(0).Control(212)=   "TXT_VPDPOLLERA_EDIT"
            Tab(0).Control(212).Enabled=   0   'False
            Tab(0).Control(213)=   "TXTCINTILLALINER_EDIT"
            Tab(0).Control(213).Enabled=   0   'False
            Tab(0).Control(214)=   "Frame3(18)"
            Tab(0).Control(214).Enabled=   0   'False
            Tab(0).Control(215)=   "Command29"
            Tab(0).Control(215).Enabled=   0   'False
            Tab(0).Control(216)=   "Command47"
            Tab(0).Control(216).Enabled=   0   'False
            Tab(0).Control(217)=   "Combo2"
            Tab(0).Control(217).Enabled=   0   'False
            Tab(0).Control(218)=   "Command43"
            Tab(0).Control(218).Enabled=   0   'False
            Tab(0).Control(219)=   "Command45"
            Tab(0).Control(219).Enabled=   0   'False
            Tab(0).Control(220)=   "TXTFUELLELINERCU"
            Tab(0).Control(220).Enabled=   0   'False
            Tab(0).Control(221)=   "TXTESLINGAMED2EDITABLE"
            Tab(0).Control(221).Enabled=   0   'False
            Tab(0).Control(222)=   "Command50"
            Tab(0).Control(222).Enabled=   0   'False
            Tab(0).ControlCount=   223
            TabCaption(1)   =   "Aprobación de Ficha Técnica"
            TabPicture(1)   =   "COTICONT.frx":0285
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Frame5(0)"
            Tab(1).Control(1)=   "Frame5(4)"
            Tab(1).ControlCount=   2
            TabCaption(2)   =   "Cotización (1)"
            TabPicture(2)   =   "COTICONT.frx":02A1
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "Frame11"
            Tab(2).Control(0).Enabled=   0   'False
            Tab(2).Control(1)=   "TXTLABEL(10)"
            Tab(2).Control(1).Enabled=   0   'False
            Tab(2).Control(2)=   "TXTNUMCOTIZACION"
            Tab(2).Control(2).Enabled=   0   'False
            Tab(2).ControlCount=   3
            TabCaption(3)   =   "Cotización (2)"
            TabPicture(3)   =   "COTICONT.frx":02BD
            Tab(3).ControlEnabled=   -1  'True
            Tab(3).Control(0)=   "Frame10"
            Tab(3).Control(0).Enabled=   0   'False
            Tab(3).ControlCount=   1
            TabCaption(4)   =   "Ficha Técnica Bolsas"
            TabPicture(4)   =   "COTICONT.frx":02D9
            Tab(4).ControlEnabled=   0   'False
            Tab(4).ControlCount=   0
            Begin VB.CommandButton Command50 
               Caption         =   "+ Datos Técnicos"
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
               Left            =   -69240
               TabIndex        =   325
               Top             =   1800
               Width           =   2000
            End
            Begin VB.TextBox TXTESLINGAMED2EDITABLE 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
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
               Left            =   -71520
               TabIndex        =   322
               Top             =   3870
               Width           =   2120
            End
            Begin VB.TextBox TXTFUELLELINERCU 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
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
               Left            =   -71520
               TabIndex        =   313
               Top             =   10110
               Width           =   2140
            End
            Begin VB.CommandButton Command45 
               Caption         =   "Ver"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   -68160
               TabIndex        =   312
               Top             =   990
               Width           =   850
            End
            Begin VB.CommandButton Command43 
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
               Height          =   285
               Left            =   -67280
               TabIndex        =   311
               Top             =   990
               Width           =   850
            End
            Begin VB.ComboBox Combo2 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   -73680
               Style           =   2  'Dropdown List
               TabIndex        =   310
               Top             =   990
               Width           =   5565
            End
            Begin VB.CommandButton Command47 
               Caption         =   "Agregar"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   -74640
               TabIndex        =   309
               Top             =   960
               Width           =   850
            End
            Begin VB.CommandButton Command29 
               Caption         =   "."
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
               Left            =   -60810
               TabIndex        =   184
               Top             =   7500
               Visible         =   0   'False
               Width           =   175
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H80000007&
               Height          =   9820
               Index           =   18
               Left            =   -67160
               TabIndex        =   275
               Top             =   1290
               Width           =   60
            End
            Begin VB.TextBox TXTCINTILLALINER_EDIT 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
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
               Left            =   -64920
               TabIndex        =   307
               Top             =   7500
               Width           =   2160
            End
            Begin VB.TextBox TXT_VPDPOLLERA_EDIT 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
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
               Left            =   -64920
               TabIndex        =   306
               ToolTipText     =   "20 A 35"
               Top             =   7185
               Width           =   2160
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H80000007&
               Height          =   52
               Index           =   5
               Left            =   -74640
               TabIndex        =   194
               Top             =   7160
               Width           =   14420
            End
            Begin VB.TextBox TXT_VPDVDESCSEGCM_EDIT 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
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
               Left            =   -71520
               TabIndex        =   305
               Top             =   7500
               Width           =   1080
            End
            Begin VB.TextBox TXT_VPDVDESCARGACM_EDIT 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
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
               Left            =   -71520
               TabIndex        =   304
               ToolTipText     =   "20 A 35"
               Top             =   7185
               Width           =   1080
            End
            Begin VB.TextBox TXT_VPDvcsegcm_EDIT 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
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
               Left            =   -68040
               TabIndex        =   303
               Top             =   7500
               Width           =   920
            End
            Begin VB.TextBox TXT_VPDCargaCm_EDIT 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
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
               Left            =   -68040
               TabIndex        =   302
               ToolTipText     =   "20 A 35"
               Top             =   7185
               Width           =   920
            End
            Begin VB.TextBox TXT_VPDCargaCm 
               Appearance      =   0  'Flat
               BackColor       =   &H80000002&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   -67125
               Locked          =   -1  'True
               TabIndex        =   181
               TabStop         =   0   'False
               Top             =   7170
               Width           =   900
            End
            Begin VB.TextBox TXT_VPDvcsegcm 
               Appearance      =   0  'Flat
               BackColor       =   &H80000002&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   -67125
               Locked          =   -1  'True
               TabIndex        =   180
               TabStop         =   0   'False
               Top             =   7500
               Width           =   900
            End
            Begin VB.TextBox Text20 
               Appearance      =   0  'Flat
               BackColor       =   &H00808080&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   -69360
               Locked          =   -1  'True
               TabIndex        =   179
               TabStop         =   0   'False
               Text            =   "V C. Seg (Cm)"
               Top             =   7500
               Width           =   1330
            End
            Begin VB.TextBox Text23 
               Appearance      =   0  'Flat
               BackColor       =   &H00808080&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   -69360
               Locked          =   -1  'True
               TabIndex        =   166
               TabStop         =   0   'False
               Text            =   "V Carga (Cm)"
               Top             =   7170
               Width           =   1330
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H80000007&
               Height          =   52
               Index           =   0
               Left            =   -74640
               TabIndex        =   193
               Top             =   8490
               Width           =   14420
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H00000000&
               Height          =   52
               Index           =   3
               Left            =   -74640
               TabIndex        =   91
               Top             =   2250
               Width           =   14420
            End
            Begin VB.TextBox Text7 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
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
               Left            =   -66240
               TabIndex        =   300
               Top             =   1950
               Width           =   420
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H80000007&
               Height          =   52
               Index           =   24
               Left            =   -74640
               TabIndex        =   299
               Top             =   10785
               Width           =   14420
            End
            Begin VB.Frame Frame10 
               Height          =   12200
               Left            =   0
               TabIndex        =   298
               Top             =   705
               Width           =   15100
            End
            Begin VB.CommandButton Command4 
               Caption         =   "Solicitar Aprobación"
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
               Left            =   -69240
               TabIndex        =   297
               Top             =   1425
               Width           =   2000
            End
            Begin VB.CommandButton Command52 
               Caption         =   "."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   -65820
               TabIndex        =   295
               Top             =   1950
               Width           =   175
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H80000007&
               Height          =   9820
               Index           =   17
               Left            =   -69420
               TabIndex        =   274
               Top             =   1290
               Width           =   60
            End
            Begin VB.TextBox TXTNUMCOTIZACION 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   -64530
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   293
               Top             =   395
               Width           =   765
            End
            Begin VB.TextBox TXTLABEL 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   10
               Left            =   -64830
               Locked          =   -1  'True
               TabIndex        =   292
               TabStop         =   0   'False
               Text            =   "N°"
               Top             =   395
               Width           =   315
            End
            Begin VB.Frame Frame11 
               Height          =   10305
               Left            =   -74880
               TabIndex        =   282
               Top             =   705
               Width           =   14895
               Begin VB.CommandButton Command49 
                  Caption         =   "Refrescar Grillas"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   720
                  TabIndex        =   323
                  Top             =   120
                  Width           =   2055
               End
               Begin VB.CommandButton Command51 
                  Caption         =   "Refrescar Cálculo de Corte"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   12240
                  TabIndex        =   294
                  Top             =   135
                  Width           =   2535
               End
               Begin VB.CommandButton Command46 
                  Caption         =   "Guardar Cotización"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   9000
                  TabIndex        =   291
                  Top             =   3000
                  Visible         =   0   'False
                  Width           =   1935
               End
               Begin VB.TextBox TXTLABEL 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0C0C0&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Index           =   93
                  Left            =   11160
                  Locked          =   -1  'True
                  TabIndex        =   284
                  TabStop         =   0   'False
                  Text            =   "SUBOTOTALES     U$S"
                  Top             =   9840
                  Width           =   1800
               End
               Begin VB.TextBox TXTSTOTSOLAPACOTIZACION 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00E0E0E0&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   12960
                  Locked          =   -1  'True
                  TabIndex        =   283
                  TabStop         =   0   'False
                  Top             =   9840
                  Width           =   1800
               End
               Begin MSFlexGridLib.MSFlexGrid MSF1 
                  Height          =   2745
                  Left            =   75
                  TabIndex        =   285
                  Top             =   480
                  Width           =   14715
                  _ExtentX        =   25956
                  _ExtentY        =   4842
                  _Version        =   393216
                  Cols            =   4
                  FixedCols       =   2
                  BackColor       =   16777215
                  BackColorFixed  =   16501405
                  BackColorBkg    =   12632256
                  GridColor       =   0
                  FillStyle       =   1
                  GridLinesFixed  =   1
                  SelectionMode   =   1
                  Appearance      =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin MSFlexGridLib.MSFlexGrid MSF2 
                  Height          =   2250
                  Left            =   75
                  TabIndex        =   286
                  Top             =   3210
                  Width           =   14715
                  _ExtentX        =   25956
                  _ExtentY        =   3969
                  _Version        =   393216
                  Cols            =   4
                  FixedCols       =   2
                  BackColor       =   16777215
                  BackColorFixed  =   16501405
                  BackColorBkg    =   12632256
                  GridColor       =   0
                  FillStyle       =   1
                  GridLinesFixed  =   1
                  SelectionMode   =   1
                  Appearance      =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin MSFlexGridLib.MSFlexGrid MSF3 
                  Height          =   1620
                  Left            =   75
                  TabIndex        =   287
                  Top             =   5460
                  Width           =   14715
                  _ExtentX        =   25956
                  _ExtentY        =   2858
                  _Version        =   393216
                  Cols            =   4
                  FixedCols       =   2
                  BackColor       =   16777215
                  BackColorFixed  =   16501405
                  BackColorBkg    =   12632256
                  GridColor       =   0
                  FillStyle       =   1
                  GridLinesFixed  =   1
                  SelectionMode   =   1
                  Appearance      =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin MSFlexGridLib.MSFlexGrid MSF5 
                  Height          =   1305
                  Left            =   75
                  TabIndex        =   288
                  Top             =   7080
                  Width           =   14715
                  _ExtentX        =   25956
                  _ExtentY        =   2302
                  _Version        =   393216
                  Cols            =   4
                  FixedCols       =   2
                  BackColor       =   16777215
                  BackColorFixed  =   16501405
                  BackColorBkg    =   12632256
                  GridColor       =   0
                  FillStyle       =   1
                  GridLinesFixed  =   1
                  SelectionMode   =   1
                  Appearance      =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin MSFlexGridLib.MSFlexGrid MSF4 
                  Height          =   1305
                  Left            =   75
                  TabIndex        =   289
                  Top             =   8385
                  Width           =   14715
                  _ExtentX        =   25956
                  _ExtentY        =   2302
                  _Version        =   393216
                  Cols            =   4
                  FixedCols       =   2
                  BackColor       =   16777215
                  BackColorFixed  =   16501405
                  BackColorBkg    =   12632256
                  GridColor       =   0
                  FillStyle       =   1
                  GridLinesFixed  =   1
                  SelectionMode   =   1
                  Appearance      =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin VB.Label Label11 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FFFFFF&
                  BackStyle       =   0  'Transparent
                  Caption         =   "RECORDAR DE ACTUALIZAR LOS PRECIOS  CUANDO ES UNA COTIZACION DE UN ITEM YA COTIZADO ANTERIORMENTE"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   492
                  Left            =   1440
                  TabIndex        =   321
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   5748
               End
               Begin VB.Label LBLNROSOLICITUDCOTIZACION 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0C0C0&
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
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Left            =   120
                  TabIndex        =   308
                  ToolTipText     =   "ID de cotización"
                  Top             =   120
                  Width           =   585
               End
               Begin VB.Label lblNuevaCotizacion 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FFFFFF&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Nueva Cotización"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   255
                  Left            =   8160
                  TabIndex        =   290
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   2505
               End
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H80000007&
               Height          =   9820
               Index           =   23
               Left            =   -74640
               TabIndex        =   280
               Top             =   1290
               Width           =   70
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H80000007&
               Height          =   9820
               Index           =   22
               Left            =   -73730
               TabIndex        =   279
               Top             =   1290
               Width           =   60
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H80000007&
               Height          =   9820
               Index           =   21
               Left            =   -60240
               TabIndex        =   278
               Top             =   1290
               Width           =   70
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H80000007&
               Height          =   8220
               Index           =   20
               Left            =   -62800
               TabIndex        =   277
               Top             =   2570
               Width           =   60
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H80000007&
               Height          =   8800
               Index           =   19
               Left            =   -65000
               TabIndex        =   276
               Top             =   2250
               Width           =   60
            End
            Begin VB.CommandButton Command36 
               Caption         =   "."
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
               Left            =   -69570
               TabIndex        =   127
               Top             =   4190
               Width           =   175
            End
            Begin VB.CommandButton Command37 
               Caption         =   "."
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
               Left            =   -69570
               TabIndex        =   122
               Top             =   4830
               Width           =   175
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H80000007&
               Height          =   9820
               Index           =   16
               Left            =   -71550
               TabIndex        =   273
               Top             =   1290
               Width           =   60
            End
            Begin VB.TextBox TXT_VPDLARGOMAMAPAROCM 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
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
               Left            =   -71520
               Locked          =   -1  'True
               TabIndex        =   207
               TabStop         =   0   'False
               Top             =   8820
               Width           =   2140
            End
            Begin VB.TextBox TXTCOMPARTCOMPARTIMENTADO 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
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
               Left            =   -71535
               Locked          =   -1  'True
               TabIndex        =   206
               TabStop         =   0   'False
               Top             =   8490
               Width           =   1950
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H80000007&
               Height          =   52
               Index           =   11
               Left            =   -74640
               TabIndex        =   236
               Top             =   5130
               Width           =   14420
            End
            Begin VB.TextBox TXTVALVVCLARGOCM 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
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
               Left            =   -71520
               Locked          =   -1  'True
               TabIndex        =   109
               TabStop         =   0   'False
               Top             =   4830
               Width           =   1960
            End
            Begin VB.TextBox TXTPOLLERALARGO 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
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
               Left            =   -71535
               Locked          =   -1  'True
               TabIndex        =   143
               TabStop         =   0   'False
               Top             =   5130
               Width           =   1960
            End
            Begin VB.TextBox TXTGRAMAJETAPA 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
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
               Left            =   -71535
               Locked          =   -1  'True
               TabIndex        =   146
               TabStop         =   0   'False
               Top             =   6150
               Width           =   1950
            End
            Begin VB.TextBox TXTGRAMAJECUERPO 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
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
               Left            =   -71535
               Locked          =   -1  'True
               TabIndex        =   144
               TabStop         =   0   'False
               Top             =   5475
               Width           =   1960
            End
            Begin VB.TextBox TXTVALVVCDIAMETROCM 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
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
               Left            =   -71520
               Locked          =   -1  'True
               TabIndex        =   110
               TabStop         =   0   'False
               Top             =   4200
               Width           =   1950
            End
            Begin VB.TextBox TXTVCLARGOCM 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
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
               Left            =   -71520
               TabIndex        =   11
               ToolTipText     =   "SI NO HAY VC: VACIO SI NO LLENAR LARGO (CM)"
               Top             =   4530
               Width           =   2160
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H80000007&
               Height          =   52
               Index           =   2
               Left            =   -74640
               TabIndex        =   77
               Top             =   3210
               Width           =   14420
            End
            Begin VB.Frame Frame5 
               Height          =   6975
               Index           =   0
               Left            =   -74880
               TabIndex        =   243
               Top             =   1050
               Width           =   11055
               Begin VB.Frame Frame5 
                  Caption         =   "NOVEDADES"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1950
                  Index           =   2
                  Left            =   240
                  TabIndex        =   251
                  Top             =   4740
                  Width           =   10620
                  Begin VB.Frame Frame5 
                     Height          =   495
                     Index           =   3
                     Left            =   0
                     TabIndex        =   254
                     Top             =   230
                     Width           =   10600
                     Begin VB.Label Label2 
                        BackStyle       =   0  'Transparent
                        Caption         =   "Usuario"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   9
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   255
                        Index           =   2
                        Left            =   8400
                        TabIndex        =   257
                        Top             =   240
                        Width           =   975
                     End
                     Begin VB.Label Label2 
                        BackStyle       =   0  'Transparent
                        Caption         =   "Detalle"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   9
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   255
                        Index           =   1
                        Left            =   1200
                        TabIndex        =   256
                        Top             =   240
                        Width           =   975
                     End
                     Begin VB.Label Label2 
                        BackStyle       =   0  'Transparent
                        Caption         =   "Fecha"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   9
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   255
                        Index           =   0
                        Left            =   120
                        TabIndex        =   255
                        Top             =   240
                        Width           =   975
                     End
                  End
                  Begin VB.TextBox Text6 
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
                     Height          =   1140
                     Left            =   1155
                     MultiLine       =   -1  'True
                     ScrollBars      =   2  'Vertical
                     TabIndex        =   253
                     Top             =   735
                     Width           =   7155
                  End
                  Begin VB.CommandButton Command41 
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
                     Height          =   745
                     Left            =   8360
                     Style           =   1  'Graphical
                     TabIndex        =   252
                     Top             =   1130
                     Width           =   750
                  End
                  Begin VB.Label Label19 
                     Appearance      =   0  'Flat
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
                     ForeColor       =   &H80000008&
                     Height          =   825
                     Left            =   0
                     TabIndex        =   261
                     Top             =   1050
                     Width           =   1170
                  End
                  Begin VB.Label Label12 
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
                     Left            =   8340
                     TabIndex        =   260
                     Top             =   840
                     Width           =   2220
                  End
                  Begin VB.Label Label6 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Registrar Anotación"
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
                     Left            =   9345
                     TabIndex        =   259
                     Top             =   1310
                     Width           =   960
                  End
                  Begin VB.Label Label9 
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
                     Height          =   345
                     Left            =   0
                     TabIndex        =   258
                     Top             =   720
                     Width           =   1170
                  End
               End
               Begin VB.Frame Frame5 
                  Caption         =   "ANOTACIONES"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   4450
                  Index           =   1
                  Left            =   240
                  TabIndex        =   244
                  Top             =   240
                  Width           =   10620
                  Begin VB.Frame Frame3 
                     Height          =   3975
                     Index           =   15
                     Left            =   8040
                     TabIndex        =   272
                     Top             =   240
                     Width           =   15
                  End
                  Begin VB.Frame Frame3 
                     Height          =   3975
                     Index           =   14
                     Left            =   1120
                     TabIndex        =   250
                     Top             =   360
                     Width           =   15
                  End
                  Begin VB.Frame Frame7 
                     Height          =   495
                     Left            =   0
                     TabIndex        =   246
                     Top             =   240
                     Width           =   10600
                     Begin VB.Label Label2 
                        BackStyle       =   0  'Transparent
                        Caption         =   "Fecha"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   9
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   255
                        Index           =   5
                        Left            =   120
                        TabIndex        =   249
                        Top             =   240
                        Width           =   975
                     End
                     Begin VB.Label Label2 
                        BackStyle       =   0  'Transparent
                        Caption         =   "Detalle"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   9
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   255
                        Index           =   4
                        Left            =   1200
                        TabIndex        =   248
                        Top             =   240
                        Width           =   975
                     End
                     Begin VB.Label Label2 
                        BackStyle       =   0  'Transparent
                        Caption         =   "Usuario"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   9
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   255
                        Index           =   3
                        Left            =   8160
                        TabIndex        =   247
                        Top             =   240
                        Width           =   975
                     End
                  End
                  Begin VB.ListBox List1 
                     Appearance      =   0  'Flat
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
                     Height          =   3624
                     Left            =   15
                     TabIndex        =   245
                     Top             =   735
                     Width           =   10575
                  End
               End
            End
            Begin VB.Frame Frame5 
               Height          =   3375
               Index           =   4
               Left            =   -63680
               TabIndex        =   239
               Top             =   1050
               Width           =   3495
               Begin VB.CommandButton Command42 
                  Caption         =   "Aprobación de F.Técnica"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   450
                  Left            =   195
                  TabIndex        =   271
                  Top             =   2760
                  Width           =   3020
               End
               Begin VB.TextBox TXTLABEL1 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0C0C0&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   280
                  Index           =   6
                  Left            =   195
                  Locked          =   -1  'True
                  TabIndex        =   270
                  TabStop         =   0   'False
                  Text            =   "FECHA"
                  Top             =   1440
                  Width           =   1000
               End
               Begin VB.TextBox TXTLABEL1 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0C0C0&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   280
                  Index           =   5
                  Left            =   195
                  Locked          =   -1  'True
                  TabIndex        =   269
                  TabStop         =   0   'False
                  Text            =   "FECHA"
                  Top             =   600
                  Width           =   1000
               End
               Begin VB.CommandButton Command40 
                  Caption         =   "Seleccionar F.Técnicas Pendientes de Aprobación"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   690
                  Left            =   195
                  TabIndex        =   268
                  Top             =   1800
                  Width           =   3020
               End
               Begin VB.TextBox TXTFECHPREPARO 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00E0E0E0&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   280
                  Left            =   1215
                  Locked          =   -1  'True
                  TabIndex        =   267
                  TabStop         =   0   'False
                  Top             =   600
                  Width           =   1000
               End
               Begin VB.TextBox TXTFECHAPROBO 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00E0E0E0&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   280
                  Left            =   1215
                  Locked          =   -1  'True
                  TabIndex        =   266
                  TabStop         =   0   'False
                  Top             =   1440
                  Width           =   1000
               End
               Begin VB.TextBox TXTPREPARO 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00E0E0E0&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   280
                  Left            =   1215
                  Locked          =   -1  'True
                  TabIndex        =   265
                  TabStop         =   0   'False
                  Top             =   240
                  Width           =   2000
               End
               Begin VB.TextBox TXTAPROBO 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00E0E0E0&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   280
                  Left            =   1215
                  Locked          =   -1  'True
                  TabIndex        =   264
                  TabStop         =   0   'False
                  ToolTipText     =   "La Aprobación Actualiza la Estructura del Item."
                  Top             =   1080
                  Width           =   2000
               End
               Begin VB.TextBox TXTLABEL1 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0C0C0&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   280
                  Index           =   31
                  Left            =   195
                  Locked          =   -1  'True
                  TabIndex        =   263
                  TabStop         =   0   'False
                  Text            =   "APROBÓ"
                  Top             =   1080
                  Width           =   1000
               End
               Begin VB.TextBox TXTLABEL1 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0C0C0&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   280
                  Index           =   35
                  Left            =   195
                  Locked          =   -1  'True
                  TabIndex        =   262
                  TabStop         =   0   'False
                  Text            =   "PREPARÓ"
                  Top             =   240
                  Width           =   1000
               End
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H80000007&
               Height          =   60
               Index           =   13
               Left            =   -74640
               TabIndex        =   238
               Top             =   1290
               Width           =   14420
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H80000007&
               Height          =   52
               Index           =   1
               Left            =   -74640
               TabIndex        =   235
               Top             =   4170
               Width           =   14420
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00808080&
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
               Index           =   38
               Left            =   -73695
               Locked          =   -1  'True
               TabIndex        =   234
               TabStop         =   0   'False
               Text            =   "Detalles de la etiqueta"
               Top             =   10440
               Width           =   2200
            End
            Begin VB.TextBox TXTDETALLEETIQUETA 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
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
               Left            =   -71520
               Locked          =   -1  'True
               TabIndex        =   233
               TabStop         =   0   'False
               Top             =   10440
               Width           =   1950
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00808080&
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
               Index           =   83
               Left            =   -67140
               Locked          =   -1  'True
               TabIndex        =   232
               TabStop         =   0   'False
               Text            =   "Detalle Etipueta"
               Top             =   10440
               Width           =   2200
            End
            Begin VB.CommandButton Command19 
               Caption         =   "."
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
               Left            =   -69555
               TabIndex        =   231
               Top             =   10440
               Width           =   175
            End
            Begin VB.CommandButton Command31 
               Caption         =   "."
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
               Left            =   -62955
               TabIndex        =   230
               Top             =   10440
               Width           =   175
            End
            Begin VB.TextBox TXTDETAETIQUETA2 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
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
               Left            =   -64935
               Locked          =   -1  'True
               TabIndex        =   229
               TabStop         =   0   'False
               Top             =   10440
               Width           =   1950
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H80000007&
               Height          =   52
               Index           =   7
               Left            =   -74640
               TabIndex        =   228
               Top             =   9450
               Width           =   14420
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H80000007&
               Height          =   52
               Index           =   10
               Left            =   -74640
               TabIndex        =   227
               Top             =   10440
               Width           =   14420
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
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
               Index           =   35
               Left            =   -73695
               Locked          =   -1  'True
               TabIndex        =   226
               TabStop         =   0   'False
               Text            =   "Liner"
               Top             =   9450
               Width           =   2200
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
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
               Index           =   36
               Left            =   -73695
               Locked          =   -1  'True
               TabIndex        =   225
               TabStop         =   0   'False
               Text            =   "Tipo "
               Top             =   9780
               Width           =   2200
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
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
               Index           =   37
               Left            =   -73695
               Locked          =   -1  'True
               TabIndex        =   224
               TabStop         =   0   'False
               Text            =   "Fuelle (Liner) c/u"
               Top             =   10110
               Width           =   2200
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   990
               Index           =   51
               Left            =   -74640
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   223
               TabStop         =   0   'False
               Text            =   "COTICONT.frx":02F5
               Top             =   9450
               Width           =   950
            End
            Begin VB.TextBox TXT_VPDFUELLELINERCU 
               Appearance      =   0  'Flat
               BackColor       =   &H80000002&
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
               Left            =   -69360
               Locked          =   -1  'True
               TabIndex        =   222
               TabStop         =   0   'False
               Top             =   10110
               Width           =   2220
            End
            Begin VB.TextBox TXTLINERLINER 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
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
               Left            =   -71520
               Locked          =   -1  'True
               TabIndex        =   221
               TabStop         =   0   'False
               Top             =   9450
               Width           =   1950
            End
            Begin VB.CommandButton Command17 
               Caption         =   "."
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
               Left            =   -69555
               TabIndex        =   220
               Top             =   9450
               Width           =   175
            End
            Begin VB.CommandButton Command18 
               Caption         =   "."
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
               Left            =   -69555
               TabIndex        =   219
               Top             =   9780
               Width           =   175
            End
            Begin VB.TextBox TXTLINERTIPO 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
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
               Left            =   -71520
               Locked          =   -1  'True
               TabIndex        =   218
               TabStop         =   0   'False
               Top             =   9780
               Width           =   1950
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
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
               Index           =   78
               Left            =   -67140
               Locked          =   -1  'True
               TabIndex        =   217
               TabStop         =   0   'False
               Text            =   "Alto (Liner)"
               Top             =   9450
               Width           =   2200
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
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
               Index           =   79
               Left            =   -67140
               Locked          =   -1  'True
               TabIndex        =   216
               TabStop         =   0   'False
               Text            =   "Micronaje (Liner)"
               Top             =   9780
               Width           =   2200
            End
            Begin VB.TextBox Text54 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
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
               Left            =   -64950
               TabIndex        =   215
               Top             =   10110
               Width           =   2160
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
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
               Index           =   80
               Left            =   -67140
               Locked          =   -1  'True
               TabIndex        =   214
               TabStop         =   0   'False
               Text            =   "Ancho"
               Top             =   10110
               Width           =   2200
            End
            Begin VB.TextBox TXT_VPDALTOLINER 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000002&
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
               Left            =   -62805
               Locked          =   -1  'True
               TabIndex        =   213
               TabStop         =   0   'False
               Top             =   9450
               Width           =   2560
            End
            Begin VB.TextBox Text56 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
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
               Left            =   -64950
               TabIndex        =   212
               Top             =   9450
               Width           =   2160
            End
            Begin VB.CommandButton Command30 
               Caption         =   "."
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
               Left            =   -62970
               TabIndex        =   211
               Top             =   9780
               Visible         =   0   'False
               Width           =   175
            End
            Begin VB.TextBox TXTMICRONAJELINERANTERIOR 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000002&
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
               Left            =   -62775
               Locked          =   -1  'True
               TabIndex        =   210
               TabStop         =   0   'False
               Text            =   "60 - 80 - 100"
               Top             =   9780
               Width           =   2550
            End
            Begin VB.TextBox TXT_VPDANCHO 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000002&
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
               Left            =   -62805
               Locked          =   -1  'True
               TabIndex        =   209
               TabStop         =   0   'False
               Top             =   10125
               Width           =   2560
            End
            Begin VB.TextBox TXT_VPDDISTANCIABOCAFONDOCM 
               Appearance      =   0  'Flat
               BackColor       =   &H80000002&
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
               Left            =   -69360
               Locked          =   -1  'True
               TabIndex        =   208
               TabStop         =   0   'False
               Top             =   9120
               Width           =   2220
            End
            Begin VB.TextBox TXT_VPDTELAMAMPARO 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
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
               Left            =   -64965
               Locked          =   -1  'True
               TabIndex        =   205
               TabStop         =   0   'False
               Top             =   9150
               Width           =   2160
            End
            Begin VB.TextBox TXT_VPDDIVISIONLATERALES 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
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
               Left            =   -64965
               Locked          =   -1  'True
               TabIndex        =   204
               TabStop         =   0   'False
               Top             =   8490
               Width           =   2160
            End
            Begin VB.TextBox TXT_VPDANCHOUTILMAMPARO 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
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
               Left            =   -64965
               Locked          =   -1  'True
               TabIndex        =   203
               TabStop         =   0   'False
               Top             =   8820
               Width           =   2160
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00808080&
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
               Index           =   77
               Left            =   -67155
               Locked          =   -1  'True
               TabIndex        =   202
               TabStop         =   0   'False
               Text            =   "Tela Mamparo"
               Top             =   9150
               Width           =   2200
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00808080&
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
               Index           =   76
               Left            =   -67155
               Locked          =   -1  'True
               TabIndex        =   201
               TabStop         =   0   'False
               Text            =   "División de Laterales"
               Top             =   8490
               Width           =   2200
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00808080&
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
               Index           =   75
               Left            =   -67155
               Locked          =   -1  'True
               TabIndex        =   200
               TabStop         =   0   'False
               Text            =   "Ancho Útil del Mamparo"
               Top             =   8820
               Width           =   2200
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00808080&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   990
               Index           =   50
               Left            =   -74640
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   199
               TabStop         =   0   'False
               Text            =   "COTICONT.frx":02FD
               Top             =   8490
               Width           =   950
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00808080&
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
               Index           =   34
               Left            =   -73695
               Locked          =   -1  'True
               TabIndex        =   198
               TabStop         =   0   'False
               Text            =   "Dist. Boca y Fdo. (cm)"
               Top             =   9150
               Width           =   2200
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00808080&
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
               Index           =   30
               Left            =   -73695
               Locked          =   -1  'True
               TabIndex        =   196
               TabStop         =   0   'False
               Text            =   "Compartimentado"
               Top             =   8490
               Width           =   2200
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H80000007&
               Height          =   52
               Index           =   6
               Left            =   -74640
               TabIndex        =   195
               Top             =   7845
               Width           =   14420
            End
            Begin VB.TextBox TXTPORTADOCUMEDIDA 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
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
               Left            =   -64935
               Locked          =   -1  'True
               TabIndex        =   192
               TabStop         =   0   'False
               Top             =   8175
               Width           =   1950
            End
            Begin VB.TextBox TXTDOCUMENTOUNIDAD 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
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
               Left            =   -64935
               Locked          =   -1  'True
               TabIndex        =   191
               TabStop         =   0   'False
               Top             =   7845
               Width           =   1950
            End
            Begin VB.CommandButton Command35 
               Caption         =   "."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   -62970
               TabIndex        =   190
               Top             =   8145
               Width           =   175
            End
            Begin VB.CommandButton Command34 
               Caption         =   "."
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
               Left            =   -62970
               TabIndex        =   189
               Top             =   7845
               Width           =   175
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
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
               Index           =   74
               Left            =   -67140
               Locked          =   -1  'True
               TabIndex        =   188
               TabStop         =   0   'False
               Text            =   "Porta Documento Medida"
               Top             =   7845
               Width           =   2200
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
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
               Index           =   73
               Left            =   -67140
               Locked          =   -1  'True
               TabIndex        =   187
               TabStop         =   0   'False
               Text            =   "Porta Docum. Cantidad"
               Top             =   8175
               Width           =   2200
            End
            Begin VB.TextBox TXT_VPDPOLLERA 
               Appearance      =   0  'Flat
               BackColor       =   &H80000002&
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
               Left            =   -62775
               Locked          =   -1  'True
               TabIndex        =   186
               TabStop         =   0   'False
               Top             =   7170
               Width           =   2560
            End
            Begin VB.TextBox TXTCINTILLALINER 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000002&
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
               Left            =   -62775
               Locked          =   -1  'True
               TabIndex        =   185
               TabStop         =   0   'False
               Text            =   "NO     /    80"
               Top             =   7500
               Width           =   2550
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00808080&
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
               Index           =   72
               Left            =   -66250
               Locked          =   -1  'True
               TabIndex        =   183
               TabStop         =   0   'False
               Text            =   "Pollera (cm)"
               Top             =   7185
               Width           =   1300
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00808080&
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
               Index           =   71
               Left            =   -66250
               Locked          =   -1  'True
               TabIndex        =   182
               TabStop         =   0   'False
               Text            =   "Liner (cm)"
               Top             =   7500
               Width           =   1300
            End
            Begin VB.CommandButton Command15 
               Caption         =   "."
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
               Left            =   -69570
               TabIndex        =   178
               Top             =   8175
               Width           =   175
            End
            Begin VB.CommandButton Command14 
               Caption         =   "."
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
               Left            =   -69570
               TabIndex        =   177
               Top             =   7845
               Width           =   175
            End
            Begin VB.TextBox TXTPDOCPDOCUMENTOUBIC 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
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
               Left            =   -71535
               Locked          =   -1  'True
               TabIndex        =   176
               TabStop         =   0   'False
               Top             =   8175
               Width           =   1950
            End
            Begin VB.TextBox TXTPDOCPORTADOCUMENTOTIPO 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
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
               Left            =   -71535
               Locked          =   -1  'True
               TabIndex        =   175
               TabStop         =   0   'False
               Top             =   7845
               Width           =   1950
            End
            Begin VB.TextBox TXT_VPDVDESCSEGCM 
               Appearance      =   0  'Flat
               BackColor       =   &H80000002&
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
               Height          =   330
               Left            =   -70440
               Locked          =   -1  'True
               TabIndex        =   174
               TabStop         =   0   'False
               Top             =   7500
               Width           =   1060
            End
            Begin VB.TextBox TXT_VPDVDESCARGACM 
               Appearance      =   0  'Flat
               BackColor       =   &H80000002&
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
               Left            =   -70440
               Locked          =   -1  'True
               TabIndex        =   173
               TabStop         =   0   'False
               Top             =   7170
               Width           =   1060
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   660
               Index           =   49
               Left            =   -74640
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   172
               TabStop         =   0   'False
               Text            =   "COTICONT.frx":0308
               Top             =   7845
               Width           =   950
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00808080&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   660
               Index           =   48
               Left            =   -74640
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   171
               TabStop         =   0   'False
               Text            =   "COTICONT.frx":0310
               Top             =   7170
               Width           =   950
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
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
               Index           =   29
               Left            =   -73695
               Locked          =   -1  'True
               TabIndex        =   170
               TabStop         =   0   'False
               Text            =   "Porta documento (ubic.)"
               Top             =   8175
               Width           =   2200
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
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
               Index           =   28
               Left            =   -73695
               Locked          =   -1  'True
               TabIndex        =   169
               TabStop         =   0   'False
               Text            =   "Porta documento Tipo"
               Top             =   7845
               Width           =   2200
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00808080&
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
               Index           =   27
               Left            =   -73695
               Locked          =   -1  'True
               TabIndex        =   168
               TabStop         =   0   'False
               Text            =   "V Desc Seg (cm)"
               Top             =   7500
               Width           =   2200
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00808080&
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
               Index           =   26
               Left            =   -73695
               Locked          =   -1  'True
               TabIndex        =   167
               TabStop         =   0   'False
               Text            =   "V Descarga (cm)"
               Top             =   7170
               Width           =   2200
            End
            Begin VB.CommandButton Command16 
               Caption         =   "."
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
               Left            =   -69570
               TabIndex        =   165
               Top             =   8520
               Width           =   175
            End
            Begin VB.TextBox TXTLAMFONDOLAMINADO 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
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
               Left            =   -71535
               Locked          =   -1  'True
               TabIndex        =   164
               TabStop         =   0   'False
               Top             =   6810
               Width           =   1950
            End
            Begin VB.TextBox TXTLAMCUERPOLAMINADO 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
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
               Left            =   -71535
               Locked          =   -1  'True
               TabIndex        =   163
               TabStop         =   0   'False
               Top             =   6480
               Width           =   1950
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H80000007&
               Height          =   52
               Index           =   8
               Left            =   -74640
               TabIndex        =   162
               Top             =   5490
               Width           =   14420
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H80000007&
               Height          =   52
               Index           =   4
               Left            =   -74640
               TabIndex        =   161
               Top             =   6480
               Width           =   14420
            End
            Begin VB.TextBox TXT_VPDTELASVALVLAMIN 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
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
               Left            =   -64950
               Locked          =   -1  'True
               TabIndex        =   160
               TabStop         =   0   'False
               Top             =   6810
               Width           =   2160
            End
            Begin VB.CommandButton Command28 
               Caption         =   "."
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
               Left            =   -62970
               TabIndex        =   159
               Top             =   6480
               Width           =   175
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
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
               Index           =   70
               Left            =   -67140
               Locked          =   -1  'True
               TabIndex        =   158
               TabStop         =   0   'False
               Text            =   "Tapa Laminada"
               Top             =   6480
               Width           =   2200
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
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
               Index           =   69
               Left            =   -67140
               Locked          =   -1  'True
               TabIndex        =   157
               TabStop         =   0   'False
               Text            =   "Tela Válvulas Laminada"
               Top             =   6810
               Width           =   2200
            End
            Begin VB.CommandButton Command13 
               Caption         =   "."
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
               Left            =   -69570
               TabIndex        =   156
               Top             =   6480
               Width           =   175
            End
            Begin VB.CommandButton Command12 
               Caption         =   "."
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
               Left            =   -69570
               TabIndex        =   155
               Top             =   6810
               Width           =   175
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   660
               Index           =   47
               Left            =   -74640
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   154
               TabStop         =   0   'False
               Text            =   "COTICONT.frx":031B
               Top             =   6480
               Width           =   950
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
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
               Index           =   25
               Left            =   -73695
               Locked          =   -1  'True
               TabIndex        =   153
               TabStop         =   0   'False
               Text            =   "Fondo Laminado"
               Top             =   6810
               Width           =   2200
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
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
               Index           =   24
               Left            =   -73695
               Locked          =   -1  'True
               TabIndex        =   152
               TabStop         =   0   'False
               Text            =   "Cuerpo Laminado"
               Top             =   6480
               Width           =   2200
            End
            Begin VB.TextBox TXTTELALAMINADA 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
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
               Left            =   -64935
               Locked          =   -1  'True
               TabIndex        =   151
               TabStop         =   0   'False
               Top             =   6480
               Width           =   1950
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00808080&
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
               Index           =   22
               Left            =   -73695
               Locked          =   -1  'True
               TabIndex        =   150
               TabStop         =   0   'False
               Text            =   "Tapa (g/m2)"
               Top             =   6135
               Width           =   2200
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00808080&
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
               Index           =   21
               Left            =   -73695
               Locked          =   -1  'True
               TabIndex        =   149
               TabStop         =   0   'False
               Text            =   "Fondo (g/m2)"
               Top             =   5820
               Width           =   2200
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00808080&
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
               Index           =   20
               Left            =   -73695
               Locked          =   -1  'True
               TabIndex        =   148
               TabStop         =   0   'False
               Text            =   "Cuerpo (g/m2)"
               Top             =   5475
               Width           =   2200
            End
            Begin VB.TextBox TXT_VPDVCGM2 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
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
               Left            =   -64950
               Locked          =   -1  'True
               TabIndex        =   142
               TabStop         =   0   'False
               Top             =   5475
               Width           =   2160
            End
            Begin VB.TextBox TXT_VPDVDGM2 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
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
               Left            =   -64950
               Locked          =   -1  'True
               TabIndex        =   141
               TabStop         =   0   'False
               Top             =   5820
               Width           =   2160
            End
            Begin VB.TextBox TXT_VPDPOLLERAGM2 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
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
               Left            =   -64950
               Locked          =   -1  'True
               TabIndex        =   140
               TabStop         =   0   'False
               Top             =   6135
               Width           =   2160
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00808080&
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
               Index           =   68
               Left            =   -67140
               Locked          =   -1  'True
               TabIndex        =   139
               TabStop         =   0   'False
               Text            =   "VC (g/m2)"
               Top             =   5475
               Width           =   2200
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00808080&
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
               Index           =   67
               Left            =   -67140
               Locked          =   -1  'True
               TabIndex        =   138
               TabStop         =   0   'False
               Text            =   "VD (g/m2)"
               Top             =   5820
               Width           =   2200
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00808080&
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
               Index           =   66
               Left            =   -67140
               Locked          =   -1  'True
               TabIndex        =   137
               TabStop         =   0   'False
               Text            =   "Pollera (g/m2)"
               Top             =   6135
               Width           =   2200
            End
            Begin VB.TextBox TXTPOLLERALAMINADA 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
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
               Left            =   -64935
               Locked          =   -1  'True
               TabIndex        =   136
               TabStop         =   0   'False
               Top             =   5145
               Width           =   1950
            End
            Begin VB.CommandButton Command27 
               Caption         =   "."
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
               Left            =   -62970
               TabIndex        =   135
               Top             =   5145
               Width           =   175
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
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
               Index           =   65
               Left            =   -67140
               Locked          =   -1  'True
               TabIndex        =   134
               TabStop         =   0   'False
               Text            =   "Pollera (Laminada)"
               Top             =   5145
               Width           =   2200
            End
            Begin VB.CommandButton Command11 
               Caption         =   "."
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
               Left            =   -69570
               TabIndex        =   133
               Top             =   6165
               Width           =   175
            End
            Begin VB.CommandButton Command10 
               Caption         =   "."
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
               Left            =   -69570
               TabIndex        =   132
               Top             =   5820
               Width           =   175
            End
            Begin VB.CommandButton Command9 
               Caption         =   "."
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
               Left            =   -69570
               TabIndex        =   131
               Top             =   5475
               Width           =   175
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00808080&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   990
               Index           =   46
               Left            =   -74640
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   130
               TabStop         =   0   'False
               Text            =   "COTICONT.frx":0321
               Top             =   5475
               Width           =   950
            End
            Begin VB.CommandButton Command38 
               Caption         =   "."
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
               Left            =   -69570
               TabIndex        =   129
               Top             =   5145
               Width           =   175
            End
            Begin VB.TextBox Text19 
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
               Height          =   330
               Left            =   -69345
               TabIndex        =   128
               TabStop         =   0   'False
               Top             =   4800
               Width           =   2220
            End
            Begin VB.TextBox TXT_VPDVDSLARGOCM 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
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
               Left            =   -64920
               Locked          =   -1  'True
               TabIndex        =   124
               Top             =   4830
               Width           =   2140
            End
            Begin VB.TextBox TXT_VPDVDLARGO 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000002&
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
               Left            =   -62790
               Locked          =   -1  'True
               TabIndex        =   123
               TabStop         =   0   'False
               Top             =   4500
               Width           =   2560
            End
            Begin VB.TextBox TXTVDLARGOCM 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
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
               Left            =   -64920
               TabIndex        =   15
               Top             =   4500
               Width           =   2160
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00808080&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   990
               Index           =   45
               Left            =   -74640
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   120
               TabStop         =   0   'False
               Text            =   "COTICONT.frx":032B
               Top             =   4170
               Width           =   950
            End
            Begin VB.TextBox TXT_VPDVCDIAMETROCM 
               Appearance      =   0  'Flat
               BackColor       =   &H80000002&
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
               Left            =   -69345
               Locked          =   -1  'True
               TabIndex        =   119
               TabStop         =   0   'False
               Top             =   4170
               Width           =   2220
            End
            Begin VB.TextBox TXT_VPDVCLARGOCM 
               Appearance      =   0  'Flat
               BackColor       =   &H80000002&
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
               Left            =   -69345
               Locked          =   -1  'True
               TabIndex        =   118
               TabStop         =   0   'False
               Top             =   4500
               Width           =   2220
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00808080&
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
               Index           =   62
               Left            =   -67125
               Locked          =   -1  'True
               TabIndex        =   117
               TabStop         =   0   'False
               Text            =   "VD (Diámetro en cm)"
               Top             =   4170
               Width           =   2200
            End
            Begin VB.TextBox TXT_VPDVDDIAMENTRO 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000002&
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
               Left            =   -62790
               Locked          =   -1  'True
               TabIndex        =   116
               TabStop         =   0   'False
               Top             =   4170
               Width           =   2560
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00808080&
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
               Index           =   63
               Left            =   -67125
               Locked          =   -1  'True
               TabIndex        =   115
               TabStop         =   0   'False
               Text            =   "VD (Largo en cm)"
               Top             =   4500
               Width           =   2200
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00808080&
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
               Index           =   64
               Left            =   -67125
               Locked          =   -1  'True
               TabIndex        =   114
               TabStop         =   0   'False
               Text            =   "VDS (Largo en cm)"
               Top             =   4830
               Width           =   2200
            End
            Begin VB.TextBox TXTVALVVDDIAMETROCM 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
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
               Left            =   -64920
               Locked          =   -1  'True
               TabIndex        =   113
               TabStop         =   0   'False
               Top             =   4170
               Width           =   1950
            End
            Begin VB.CommandButton Command26 
               Caption         =   "."
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
               Left            =   -62970
               TabIndex        =   112
               Top             =   4170
               Width           =   175
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H00000000&
               Height          =   52
               Index           =   12
               Left            =   -74640
               TabIndex        =   111
               Top             =   4185
               Width           =   14420
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   930
               Index           =   44
               Left            =   -74640
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   105
               TabStop         =   0   'False
               Text            =   "COTICONT.frx":0336
               Top             =   3240
               Width           =   950
            End
            Begin VB.TextBox TXTESLINGAOJOCM 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
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
               Left            =   -71520
               TabIndex        =   9
               ToolTipText     =   "20 A 35"
               Top             =   3240
               Width           =   2160
            End
            Begin VB.TextBox TXTESLINGAMEDIDA1CM 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
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
               Left            =   -71520
               TabIndex        =   10
               Top             =   3555
               Width           =   2160
            End
            Begin VB.TextBox TXTESLINGAMEDIDA2CM 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
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
               Left            =   -69360
               Locked          =   -1  'True
               TabIndex        =   104
               TabStop         =   0   'False
               Top             =   3870
               Width           =   2220
            End
            Begin VB.TextBox TXT_VPDFESLINGA 
               Appearance      =   0  'Flat
               BackColor       =   &H80000002&
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
               Left            =   -69360
               Locked          =   -1  'True
               TabIndex        =   103
               TabStop         =   0   'False
               Top             =   3555
               Width           =   2220
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
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
               Index           =   59
               Left            =   -67140
               Locked          =   -1  'True
               TabIndex        =   102
               TabStop         =   0   'False
               Text            =   "Eslinga Perimetral"
               Top             =   3870
               Width           =   2200
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
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
               Index           =   60
               Left            =   -67140
               Locked          =   -1  'True
               TabIndex        =   101
               TabStop         =   0   'False
               Text            =   "Eslinga (Ancho Cm)"
               Top             =   3555
               Width           =   2200
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
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
               Index           =   61
               Left            =   -67140
               Locked          =   -1  'True
               TabIndex        =   100
               TabStop         =   0   'False
               Text            =   "Tipo de Costura"
               Top             =   3240
               Width           =   2200
            End
            Begin VB.TextBox TXTTIPOCOSTURA 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
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
               Left            =   -64935
               Locked          =   -1  'True
               TabIndex        =   99
               TabStop         =   0   'False
               Top             =   3240
               Width           =   1950
            End
            Begin VB.TextBox TXTESLINGAANCHO 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
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
               Left            =   -64935
               Locked          =   -1  'True
               TabIndex        =   98
               TabStop         =   0   'False
               Top             =   3555
               Width           =   1950
            End
            Begin VB.TextBox TXTESLINGAPERIMETRAL 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
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
               Left            =   -64935
               Locked          =   -1  'True
               TabIndex        =   97
               TabStop         =   0   'False
               Top             =   3870
               Width           =   1950
            End
            Begin VB.CommandButton Command23 
               Caption         =   "."
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
               Left            =   -62970
               TabIndex        =   96
               Top             =   3870
               Width           =   175
            End
            Begin VB.CommandButton Command24 
               Caption         =   "."
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
               Left            =   -62970
               TabIndex        =   95
               Top             =   3555
               Width           =   175
            End
            Begin VB.CommandButton Command25 
               Caption         =   "."
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
               Left            =   -62970
               TabIndex        =   94
               Top             =   3240
               Width           =   175
            End
            Begin VB.CommandButton Command21 
               Caption         =   "."
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
               Left            =   -62970
               TabIndex        =   93
               Top             =   2910
               Width           =   175
            End
            Begin VB.TextBox Text33 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00808080&
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
               Left            =   -64935
               Locked          =   -1  'True
               TabIndex        =   92
               TabStop         =   0   'False
               Text            =   "DOBLE CADENETA"
               Top             =   2250
               Width           =   4690
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00808080&
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
               Index           =   57
               Left            =   -67140
               Locked          =   -1  'True
               TabIndex        =   89
               TabStop         =   0   'False
               Text            =   "Costura con Pabilo"
               Top             =   2580
               Width           =   2200
            End
            Begin VB.CommandButton Command8 
               Caption         =   "."
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
               Left            =   -67320
               TabIndex        =   88
               Top             =   2910
               Width           =   175
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00808080&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   990
               Index           =   43
               Left            =   -74640
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   85
               TabStop         =   0   'False
               Text            =   "COTICONT.frx":0340
               Top             =   2250
               Width           =   950
            End
            Begin VB.TextBox TXTCUERPOALTOCM 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
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
               Left            =   -71520
               TabIndex        =   8
               Top             =   2910
               Width           =   2160
            End
            Begin VB.TextBox TXTTELAVENTILADA 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
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
               Left            =   -69360
               TabIndex        =   84
               TabStop         =   0   'False
               Top             =   2910
               Width           =   2025
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00808080&
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
               Index           =   56
               Left            =   -67140
               Locked          =   -1  'True
               TabIndex        =   83
               TabStop         =   0   'False
               Text            =   "Costura Externa"
               Top             =   2910
               Width           =   2200
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00808080&
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
               Index           =   58
               Left            =   -67140
               Locked          =   -1  'True
               TabIndex        =   82
               TabStop         =   0   'False
               Text            =   "Tipo de Costura"
               Top             =   2250
               Width           =   2200
            End
            Begin VB.TextBox TXTCOSTURACONPABILO 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
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
               Left            =   -64935
               Locked          =   -1  'True
               TabIndex        =   81
               TabStop         =   0   'False
               Top             =   2580
               Width           =   1950
            End
            Begin VB.TextBox TXTCOSTURAEXTERNA 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
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
               Left            =   -64935
               Locked          =   -1  'True
               TabIndex        =   80
               TabStop         =   0   'False
               Top             =   2910
               Width           =   1950
            End
            Begin VB.CommandButton Command22 
               Caption         =   "."
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
               Left            =   -62970
               TabIndex        =   79
               Top             =   2580
               Width           =   175
            End
            Begin VB.TextBox Text38 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
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
               Left            =   -69360
               Locked          =   -1  'True
               TabIndex        =   78
               TabStop         =   0   'False
               Text            =   "Tela Ventilada"
               Top             =   2580
               Width           =   2220
            End
            Begin VB.TextBox TXTANCHOCM 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
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
               Left            =   -71520
               TabIndex        =   6
               Top             =   2250
               Width           =   2160
            End
            Begin VB.TextBox TXTCUERPO_LARGOCM 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
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
               Left            =   -71520
               TabIndex        =   7
               Top             =   2580
               Width           =   2160
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
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
               Index           =   23
               Left            =   -73695
               Locked          =   -1  'True
               TabIndex        =   76
               TabStop         =   0   'False
               Text            =   "Impresión"
               Top             =   1290
               Width           =   2200
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
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
               IMEMode         =   3  'DISABLE
               Index           =   1
               Left            =   -73695
               Locked          =   -1  'True
               TabIndex        =   75
               TabStop         =   0   'False
               Text            =   "Cantidad de Caras"
               Top             =   1890
               Width           =   2200
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
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
               Index           =   42
               Left            =   -74640
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   74
               TabStop         =   0   'False
               Text            =   "COTICONT.frx":0349
               Top             =   1290
               Width           =   950
            End
            Begin VB.CommandButton Command5 
               Caption         =   "."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Left            =   -69570
               TabIndex        =   73
               Top             =   1290
               Width           =   175
            End
            Begin VB.TextBox TXTIMPRESION 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
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
               Left            =   -71535
               Locked          =   -1  'True
               TabIndex        =   16
               TabStop         =   0   'False
               ToolTipText     =   "Doble Click Para Limpiar Casillero"
               Top             =   1290
               Width           =   1950
            End
            Begin VB.TextBox TXTCANTICARAS 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
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
               Left            =   -71535
               Locked          =   -1  'True
               TabIndex        =   72
               TabStop         =   0   'False
               ToolTipText     =   "Doble Click Para Limpiar Casillero"
               Top             =   1890
               Width           =   1950
            End
            Begin VB.TextBox TXTLABEL 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
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
               Index           =   53
               Left            =   -67140
               Locked          =   -1  'True
               TabIndex        =   71
               TabStop         =   0   'False
               Text            =   "Colores"
               Top             =   1950
               Width           =   1485
            End
            Begin VB.TextBox TXTLABEL 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
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
               Index           =   54
               Left            =   -67140
               Locked          =   -1  'True
               TabIndex        =   70
               TabStop         =   0   'False
               Text            =   "Especificación Impresión de Conteplast"
               Top             =   1620
               Width           =   4605
            End
            Begin VB.TextBox TXTLABEL 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
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
               Index           =   55
               Left            =   -67140
               Locked          =   -1  'True
               TabIndex        =   69
               TabStop         =   0   'False
               Text            =   "Especificación Impresión del Cliente"
               Top             =   1290
               Width           =   4605
            End
            Begin VB.TextBox Text29 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
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
               Left            =   -62535
               TabIndex        =   12
               Top             =   1290
               Width           =   2300
            End
            Begin VB.TextBox Text31 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
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
               Left            =   -62535
               TabIndex        =   13
               Top             =   1620
               Width           =   2300
            End
            Begin VB.TextBox Text32 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
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
               Left            =   -65640
               Locked          =   -1  'True
               TabIndex        =   14
               ToolTipText     =   "Doble Click Para Limpiar Colores"
               Top             =   1950
               Width           =   5420
            End
            Begin VB.CommandButton Command6 
               Caption         =   "."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Left            =   -69570
               TabIndex        =   68
               Top             =   1890
               Width           =   175
            End
            Begin VB.TextBox TXTOBSAMPLIA 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   900
               Left            =   -74640
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   67
               Top             =   11385
               Width           =   14400
            End
            Begin VB.TextBox TXTLABEL 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
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
               Index           =   8
               Left            =   -74640
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   66
               TabStop         =   0   'False
               Text            =   "COTICONT.frx":035B
               Top             =   11145
               Width           =   1560
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H80000007&
               Height          =   52
               Index           =   9
               Left            =   -74640
               TabIndex        =   65
               Top             =   11130
               Width           =   14440
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
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
               Index           =   39
               Left            =   -73695
               Locked          =   -1  'True
               TabIndex        =   64
               TabStop         =   0   'False
               Text            =   "Embalaje U/Fardo"
               Top             =   10770
               Width           =   2200
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
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
               Index           =   52
               Left            =   -74640
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   63
               TabStop         =   0   'False
               Text            =   "COTICONT.frx":036C
               Top             =   10770
               Width           =   950
            End
            Begin VB.TextBox TXTEMBALAJEFARDO 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
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
               Left            =   -71520
               Locked          =   -1  'True
               TabIndex        =   62
               TabStop         =   0   'False
               Top             =   10770
               Width           =   1950
            End
            Begin VB.CommandButton Command20 
               Caption         =   "."
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
               Left            =   -69555
               TabIndex        =   61
               Top             =   10770
               Width           =   175
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
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
               Index           =   82
               Left            =   -67140
               Locked          =   -1  'True
               TabIndex        =   60
               TabStop         =   0   'False
               Text            =   "Pallet"
               Top             =   10770
               Width           =   2200
            End
            Begin VB.CommandButton Command32 
               Caption         =   "."
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
               Left            =   -62955
               TabIndex        =   59
               Top             =   10770
               Width           =   175
            End
            Begin VB.TextBox TXTPALLET 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
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
               Left            =   -64935
               Locked          =   -1  'True
               TabIndex        =   58
               TabStop         =   0   'False
               Top             =   10770
               Width           =   1950
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
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
               Index           =   81
               Left            =   -62805
               Locked          =   -1  'True
               TabIndex        =   57
               TabStop         =   0   'False
               Text            =   "Un/Pallet"
               Top             =   10770
               Width           =   980
            End
            Begin VB.TextBox TXTUNPALLET 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
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
               Left            =   -61830
               Locked          =   -1  'True
               TabIndex        =   56
               TabStop         =   0   'False
               Top             =   10770
               Width           =   1420
            End
            Begin VB.CommandButton Command33 
               Caption         =   "."
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
               Left            =   -60405
               TabIndex        =   55
               Top             =   10770
               Width           =   175
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00808080&
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
               Index           =   4
               Left            =   -73680
               Locked          =   -1  'True
               TabIndex        =   86
               TabStop         =   0   'False
               Text            =   "Alto (Cm)"
               Top             =   2910
               Width           =   2200
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00808080&
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
               Index           =   2
               Left            =   -73680
               Locked          =   -1  'True
               TabIndex        =   87
               TabStop         =   0   'False
               Text            =   "Ancho (Cm)"
               Top             =   2250
               Width           =   2200
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00808080&
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
               Index           =   3
               Left            =   -73680
               Locked          =   -1  'True
               TabIndex        =   90
               TabStop         =   0   'False
               Text            =   "Largo (Cm)"
               Top             =   2580
               Width           =   2200
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
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
               Index           =   7
               Left            =   -73680
               Locked          =   -1  'True
               TabIndex        =   106
               TabStop         =   0   'False
               Text            =   "Eslinga (Medida2 Cm)"
               Top             =   3870
               Width           =   2200
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
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
               Index           =   6
               Left            =   -73680
               Locked          =   -1  'True
               TabIndex        =   107
               TabStop         =   0   'False
               Text            =   "Eslinga (Medida1 Cm)"
               Top             =   3555
               Width           =   2200
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
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
               Index           =   5
               Left            =   -73680
               Locked          =   -1  'True
               TabIndex        =   108
               TabStop         =   0   'False
               Text            =   "Eslinga (Ojo Cm)"
               Top             =   3240
               Width           =   2200
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00808080&
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
               Index           =   18
               Left            =   -73680
               Locked          =   -1  'True
               TabIndex        =   126
               TabStop         =   0   'False
               Text            =   "VCS (Largo en cm)"
               Top             =   4830
               Width           =   2200
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00808080&
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
               Index           =   16
               Left            =   -73680
               Locked          =   -1  'True
               TabIndex        =   125
               TabStop         =   0   'False
               Text            =   "VC (Diámetro en cm)"
               Top             =   4170
               Width           =   2200
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00808080&
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
               Index           =   17
               Left            =   -73680
               Locked          =   -1  'True
               TabIndex        =   121
               TabStop         =   0   'False
               Text            =   "VC (Largo en cm)"
               Top             =   4500
               Width           =   2200
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
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
               Index           =   19
               Left            =   -73680
               Locked          =   -1  'True
               TabIndex        =   147
               TabStop         =   0   'False
               Text            =   "Pollera (Largo en cm)"
               Top             =   5130
               Width           =   2200
            End
            Begin VB.TextBox TXTLABEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00808080&
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
               Index           =   33
               Left            =   -73695
               Locked          =   -1  'True
               TabIndex        =   197
               TabStop         =   0   'False
               Text            =   "Largo del Mamparo (cm)"
               Top             =   8820
               Width           =   2200
            End
            Begin VB.TextBox TXTDISTBOCAFDOEIDIT 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
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
               Left            =   -71520
               TabIndex        =   296
               Top             =   9120
               Width           =   2140
            End
            Begin VB.TextBox TXTGRAMAJEFONDO 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
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
               Left            =   -71535
               Locked          =   -1  'True
               TabIndex        =   145
               TabStop         =   0   'False
               Top             =   5820
               Width           =   1960
            End
            Begin VB.TextBox TXTMICRONAJELINER 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
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
               Left            =   -64950
               TabIndex        =   301
               Top             =   9780
               Width           =   2160
            End
            Begin VB.Label Label8 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "**** Espere a Que Desaparezca esta Leyenda****      **** Se  esta actualizando la base de datos *****"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   615
               Left            =   -66360
               TabIndex        =   318
               Top             =   705
               Visible         =   0   'False
               Width           =   6135
            End
            Begin VB.Label LBLNUEVAFICHATECNICA 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Nueva Ficha Técnica"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   18
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   855
               Left            =   -69360
               TabIndex        =   237
               Top             =   5490
               Visible         =   0   'False
               Width           =   2205
            End
         End
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   7075
         Left            =   15170
         TabIndex        =   42
         Top             =   120
         Width           =   300
      End
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   495
      Left            =   7200
      TabIndex        =   317
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   495
      Left            =   7200
      TabIndex        =   316
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   7200
      TabIndex        =   315
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   7200
      X2              =   8400
      Y1              =   6960
      Y2              =   7440
   End
   Begin VB.Line Line2 
      BorderWidth     =   50
      X1              =   2760
      X2              =   8880
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   0
      TabIndex        =   281
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Menu mnuarchivo 
      Caption         =   "Transacciones"
      Begin VB.Menu mnuNuevaReceta 
         Caption         =   "Nueva Ficha Técnica"
      End
      Begin VB.Menu mnRecetaExistentegral 
         Caption         =   "Abrir FichaTécnica Existente"
         Begin VB.Menu mnRecetaExistente 
            Caption         =   "Del Vendedor Activo"
         End
         Begin VB.Menu todasAprobadasSinFiltroVend 
            Caption         =   "Solo Aprobadas Vendedor Activo"
         End
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu missolicitudesAprobacionft 
         Caption         =   "Mis Solicitudes de Aprobación de Fichas Técnicas"
      End
      Begin VB.Menu mnuGuardar 
         Caption         =   "Guardar Ficha Técnica"
      End
      Begin VB.Menu mnuModificarFTG 
         Caption         =   "Modificar Ficha Técnica Aprobada"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuAnularCotización 
         Caption         =   "Anulación de  Ficha Técnica"
      End
      Begin VB.Menu S3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGral 
         Caption         =   "Generación por Copia de F.Técnicas Existentes"
         Begin VB.Menu mnuGuardarComoNuevaFtecnica 
            Caption         =   "Copiar  F.Técnica Activa (Puede Aplicarse a Otro Cliente) "
         End
      End
      Begin VB.Menu mnuGenexcopiaCodigo 
         Caption         =   "Generación de Nueva Versión de Código Existente Mismo Cliente"
         Begin VB.Menu mnunewversxcliente 
            Caption         =   "Filtro por Cliente"
         End
         Begin VB.Menu mnuseleccodigo 
            Caption         =   "Seleccionar de lista de Códigos"
         End
      End
      Begin VB.Menu mnuimprimir 
         Caption         =   "Imprimir"
         Begin VB.Menu mnuimprimirfichaactiva 
            Caption         =   "Ficha Técnica Activa"
         End
         Begin VB.Menu mnuimprimecotizacion 
            Caption         =   "Seleccionar Cotización a Imprimir"
         End
      End
      Begin VB.Menu S2 
         Caption         =   "-"
      End
      Begin VB.Menu nuevaCotizacion 
         Caption         =   "Nueva Cotización"
         Begin VB.Menu mnunewCotizacion 
            Caption         =   "Por Selección de Código"
         End
         Begin VB.Menu mnuxSolicituddeCotizacion 
            Caption         =   "Por Solicitudes de Cotizaciones Pendientes"
         End
      End
      Begin VB.Menu mnuAbrirCotExistente 
         Caption         =   "Abrir Cotización Existente"
      End
      Begin VB.Menu mnuGuardarCoti 
         Caption         =   "Guardar Cotización"
      End
      Begin VB.Menu S1 
         Caption         =   "-"
      End
      Begin VB.Menu mnusalir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu mnuConsList 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu mnuftecnicas 
         Caption         =   "Fichas Técnicas"
      End
      Begin VB.Menu mnuCotizaciones 
         Caption         =   "Cotizaciones"
      End
      Begin VB.Menu mnuCodiRang 
         Caption         =   "Códigos Rango y Asociados"
      End
   End
   Begin VB.Menu mnuVer 
      Caption         =   "Ver"
   End
   Begin VB.Menu mnuUtilidades 
      Caption         =   "Utilidades"
      Begin VB.Menu mnuBajaFichaTecnica 
         Caption         =   "Baja de Fichas Técnicas"
      End
      Begin VB.Menu mnuBajaCotización 
         Caption         =   "Baja de Solicitudes de Cotización"
      End
      Begin VB.Menu mnuMarcaNoMostrar 
         Caption         =   "Marcar Fichas Técnicas a No Mostrar"
         Index           =   0
      End
      Begin VB.Menu mnuMarcaNoMostrar 
         Caption         =   "Revertir Fichas Técnicas a No Mostrar"
         Index           =   1
      End
      Begin VB.Menu mnuMarcaNoMostrar 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuMarcaNoMostrar 
         Caption         =   "Marcar Cotizaciones a No Mostrar"
         Index           =   3
      End
      Begin VB.Menu mnuMarcaNoMostrar 
         Caption         =   "Revertir Cotizaciones a No Mostrar"
         Index           =   4
      End
   End
   Begin VB.Menu mnuConfiguracion 
      Caption         =   "Configuración"
      Begin VB.Menu mnuTablasAuxiliares 
         Caption         =   "Tablas Auxiliares"
      End
      Begin VB.Menu mnuBB 
         Caption         =   "Configuración Big Bag"
      End
      Begin VB.Menu mnucortebigbag 
         Caption         =   "Configuración Cáculo de Corte según BigBag"
      End
      Begin VB.Menu mnuCodRango 
         Caption         =   "Configuración Códigos Rango"
      End
      Begin VB.Menu mnuNivelesAprobacion 
         Caption         =   "Nivles de Aprobación"
      End
      Begin VB.Menu mnuConfigWaldbot 
         Caption         =   "Configuración General"
      End
      Begin VB.Menu mnuSeleccionCarpetaImagen 
         Caption         =   "Seleccionar Carpeta Donde se Guardaran las Imagenes"
      End
   End
   Begin VB.Menu mnuAD 
      Caption         =   "Accesos Directos"
      Begin VB.Menu mnuAltaMP 
         Caption         =   "Alta de Materias Primas"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuPedidos 
         Caption         =   "Pedidos a Clientes"
      End
      Begin VB.Menu mnuActualizarBaseClientes 
         Caption         =   "Actualizar Base de Datos Clientes"
      End
      Begin VB.Menu mnuCostos 
         Caption         =   "Costos Analisis x Rango de Fechas"
      End
   End
End
Attribute VB_Name = "COTICONT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    Public MODIFIC%
    Public MODIFICCOT1%
    Public MODIFICCOT2%
    Dim ActualizaListaSeleccion As Boolean
    Private Declare Function GetShortPathName Lib "KERNEL32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBuffer As Long) As Long
    Dim WinRar As String, WinZip As String
Sub AGREGARRUTAIMAG(Optional FTEC&, Optional OK%)

      CD1.Filter = "(*.*)|*.*|"
      CD1.ShowOpen
      Ruta$ = CD1.FileName
      If TRIM(Ruta$) = "" Then Exit Sub
      '
20    DESCRI$ = TRIM$(InputBox$("Ingrese Texto ", "Guardar Como", CODI$))
      If DESCRI$ = "" Then Exit Sub
      DESCRI$ = UCase$(DESCRI$)
      If Len(DESCRI$) > 17 Then
        Call COMUNI("Longitud de Descripción \Supera La Cantidad Máxima de Caracteres (17)")
        GoTo 20
      End If
      DESCRI$ = DESCRI$ & "_" & FTEC&
      On Error Resume Next
      CARPETA_ORIGEN$ = Ruta$
      POS1% = InStrRev(Ruta$, "\")
      POS2% = InStrRev(Ruta$, ".")
      CARPETA_DESTINO$ = Control("CARPIMAG", "UBICIMAG") & "\" & TRIM$(DESCRI$) & Mid$(Ruta$, POS2%)
    
      CONDI$ = "NUMRECETACTP = " & FTEC&
      CONDI$ = CONDI$ + " AND Descripcion = '" & DESCRI$ & "'"
      SIYAEXISTE% = CantRegistros&("DOCURELFT", CONDI$, "NOMESS")
      If SIYAEXISTE% > 0 Then
         Call COMUNI("Archivo Existente en Carpeta\Puede Eliminar el Existente y Volver a Intentar\O Bien Guardarlo Con Otro Nombre")
         OK% = -1
         Exit Sub
      End If
      '
      FileCopy CARPETA_ORIGEN$, CARPETA_DESTINO$
      If Err > 0 Then
       A = 1
      End If
      On Error GoTo 0
      'AGREGAR INFORMACION
      Call REPLA(TX$, DESCRI$, 1, 24)
      Call REPLA(TX$, CARPETA_DESTINO$, 32, 512)
      Combo2.AddItem TX$

End Sub

Sub CARGAR_GRILLAS_COTIZACION_YA_TIENE_FORMULA(CI1%)
    'SI YA TIENE ESTRUCTURA (APROBADA) EL CODIGO A COTIZAR CARGA LA ULTIMA QUE TIENE DE ESE CODIGO
    NROCOT& = XVALO(VALOBADA("COTI1", "NROCOTIZACION", "CODINTCONJ = " & CI1% & " ORDER BY COTI1_ID DESC", "NOMESS"))
    If NROCOT& < 1 Then Exit Sub
    Dim MSF As msFlexGrid
    For i% = 1 To 5
        REG& = 0
        Select Case i%
          Case 1
            Set MSF = MSF1
            MSF1.Rows = 1
          Case 2
            Set MSF = msf2
            msf2.Rows = 2
          Case 3
            Set MSF = MSF3
            MSF3.Rows = 3
          Case 4
            Set MSF = MSF4
            MSF4.Rows = 1
          Case 5
            Set MSF = MSF5
            MSF5.Rows = 1
        End Select
          
        SQLX$ = "SELECT * FROM COTI1 WITH(NOLOCK)"
        SQLX$ = SQLX$ + " WHERE NROCOTIZACION  = " & XVALO(NROCOT&)
        SQLX$ = SQLX$ + " AND NAMEMSF = '" & MSF.Name & "' "
        SQLX$ = SQLX$ + " ORDER BY ORDEN"
        Set RS = READSET(SQLX$)
        With RS
          Do While Not .EOF
            REG& = REG& + 1
            If MSF.Rows < REG + 1 Then MSF.Rows = REG& + 1
            MSF.TextMatrix(REG&, 1) = SAFETEXT$(!COMPONENTE)
            MSF.TextMatrix(REG&, 2) = SAFETEXT$(!F_TECNICA)
            MSF.TextMatrix(REG&, 3) = SAFETEXT$(!Codigo)
            MSF.TextMatrix(REG&, 4) = SAFETEXT$(!Descripcion)
            MSF.TextMatrix(REG&, 5) = XVALO(!moduxbb)
            MSF.TextMatrix(REG&, 6) = FORMATNUM$(XVALO(!corte), "F12.2")
            MSF.TextMatrix(REG&, 7) = FORMATNUM$(XVALO(!moduxbb) * XVALO(!corte), "F12.4")
            COSTO# = COSUNI(XVALO(!CODINTELEM)) 'RECALCULO CON EL PRECIO ACTUAL EL RESTO MANTENGO TODO
            MSF.TextMatrix(REG&, 8) = FORMATNUM$(XVALO(COSTO#), "F12.4")
            If MSF.Name = "MSF1" Then 'ESCRIBE LOS DATOS DE CONTROL DEL CORTE QUE VIENEN POR FORMULA
              MSF.TextMatrix(REG&, 10) = FORMATNUM$(XVALO(!RESPACORTE), "F12.2")
            End If
            If MSF.Name = "MSF5" Then
                MSF.TextMatrix(REG&, 9) = FORMATNUM$((XVALO(MSF.TextMatrix(REG&, 6)) * XVALO(COSTO#)) / 100, "F12.2")
            Else
              MSF.TextMatrix(REG&, 9) = FORMATNUM$((XVALO(MSF.TextMatrix(REG&, 7)) * XVALO(COSTO#)) / 100, "F12.2")
            End If

            .MoveNext
          Loop
        End With
    Next i%
    RS.Close
    Set RS = Nothing


End Sub



Sub CONFIGMENU()
    'CONFIGURACIONES POR VENDEDOR SOLO SUPERVISOR VE TODO
    USNU% = USNBR% Mod 100
    If USNU% > 1 And DerechoSupervisor% < 1 Then
       VENDESEGUNUSUARIO% = LOGVEND.VendeVsUsuario(USNU%)
       If VENDESEGUNUSUARIO% > 1 Then  ' SI NO SON VENDEDORES DEBEN ASIGNARSE VENDEDOR 1 GERENCIA.
            mnuModificarFTG.Enabled = False
            nuevaCotizacion.Enabled = False
            mnuAbrirCotExistente.Enabled = False
            mnuGuardarCoti.Enabled = False
            mnuAltaMP.Enabled = False
'            mnuAD.Enabled = False
            Command2.Enabled = False
       End If
    End If
End Sub

Public Function GetShortPath(strFileName As String) As String
 Dim ret As Long, strPath As String
 strPath = String$(165, 0)
 ret = GetShortPathName(strFileName, strPath, 164)
 GetShortPath = Left(strPath, ret)
End Function


Function CORTE_CONTEPLAST(formulaTapa$, ALTO, ANCHO, Largo, PARTE$, FORMATOX$)
     'ESTA INFORMACION DE LA FORMULA VIENE DE LA TABLA CALCUCORTE
     'EJ. (;ALTO; * 2 )/2
     'VALIDO FORMULA
     formulaTapa$ = UCase$(formulaTapa$)
     VALIFORMU$ = formulaTapa$
     VALIFORMU$ = REPLACAR(VALIFORMU$, "ALTO", "")
     VALIFORMU$ = REPLACAR(VALIFORMU$, "ANCHO", "")
     VALIFORMU$ = REPLACAR(VALIFORMU$, "LARGO", "")
     
     For i% = 1 To Len(VALIFORMU$)
       CARACTER$ = Left(VALIFORMU$, 1)
       ASCIX% = Asc(CARACTER$)
       If ASCIX% < 32 Then
         CORTE_CONTEPLAST = -1
       ElseIf ASCIX% < 40 And ASCIX% > 32 Then
         CORTE_CONTEPLAST = -1
       ElseIf ASCIX% = 46 Or ASCIX% = 58 Then
          CORTE_CONTEPLAST = -1
       ElseIf ASCIX% > 59 Then
           formulaTapa$ = -1
       End If
       If Len(VALIFORMU$) > 1 Then
          VALIFORMU$ = Mid(VALIFORMU$, 2)
       Else
          Exit For
       End If
     Next i%
     If CORTE_CONTEPLAST < 0 Then
       Call COMUNI("Formula de " & PARTE$ & " Inválida: " & formulaTapa$)
       Exit Function
     End If
      '
     Dim strArray() As String
     On Error Resume Next
     strArray = Split(formulaTapa$, ";")
     Formula1$ = formulaTapa$
     'VERIFICA SI EN LA FORMULA EXISTEN LAS PALABRAS ALTO, LARGO Y ANCHO EN LOS TROZOS POR QUE TIENEN QUE ESTAR SEPARADOS POR ;
     For i% = 0 To UBound(strArray)
       Dato$ = strArray(i%)
       If Dato$ = "ALTO" Or Dato$ = "ANCHO" Or Dato$ = "LARGO" Then 'EVALUA SI EL ANCHO ALTO O LARGO DENTRO DE LA FORMULA
         Select Case Dato$
            Case "ALTO"
               Formula1$ = REPLACAR(Formula1$, "ALTO", SAFETEXT$(XVALO(ALTO)))
            Case "ANCHO"
               Formula1$ = REPLACAR(Formula1$, "ANCHO", SAFETEXT$(XVALO(ANCHO)))
            Case "LARGO"
               Formula1$ = REPLACAR(Formula1$, "LARGO", SAFETEXT$(XVALO(Largo)))
         End Select
       Else
         CORTE_CONTEPLAST = 0
       End If
     Next i%
     
     Formula1$ = REPLACAR(Formula1$, ";", "")
     CORTE_CONTEPLAST = FORMATNUM$(CALCULARFORMULAS(Formula1$), FORMATOX$)

End Function





Sub CARGAMSF1(CARGARDATOS%)
           ANCHO = XVALO(TXTANCHOCM)
           Largo = XVALO(TXTCUERPO_LARGOCM)
           ALTO = XVALO(TXTCUERPOALTOCM)
           ANCHOLARGOYALTO$ = "(" & SAFETEXT(ANCHO) & "-" & SAFETEXT(Largo) & "-" & SAFETEXT(ALTO) & ")"
'           If ANCHO <= 0 Or LARGO <= 0 Or ALTO <= 0 Then
'                 Call COMUNI("Atención Imposible Continuar Medidas de Alto, Largo o Ancho Incompletas")
'                 Exit Sub
'           End If
'VER DE PONER LAS FILAS OCULTAS CON LA FORMULA  Y QUE LUEGO LA TRADUZCA.
           Campos$ = "COMPONENTE/A24;FICHA TECNICA(g-m2)/A20;CODIGO/A16;DESCRIPCION/A24;MODULO/F12;CORTE/F10.2;CONSUMO (cm)/f12.2;PRECIO (u$s*M)/F12.2;PRECIO (u$s)/F12.2;CONTROLCORTE/A0"
           Call CARGA_ENCABEZADO(MSF1, Campos$, Me)
           MSF1.Rows = 9
           '********************no cambiar estos textos estan tenidos en cuenta cuando se graba y para los frm *******************************
           MSF1.TextMatrix(1, 1) = "TAPA " & ANCHOLARGOYALTO$
           MSF1.TextMatrix(2, 1) = "POLLERA " & ANCHOLARGOYALTO$
           MSF1.TextMatrix(3, 1) = "CUERPO " & ANCHOLARGOYALTO$
           MSF1.TextMatrix(4, 1) = "LATERALES " & ANCHOLARGOYALTO$
           MSF1.TextMatrix(5, 1) = "FONDO " & ANCHOLARGOYALTO$
           MSF1.TextMatrix(6, 1) = "UPANEL " & ANCHOLARGOYALTO$
           MSF1.TextMatrix(7, 1) = "VÁLVULAS DE CARGA [CM]"
           MSF1.TextMatrix(8, 1) = "VÁLVULAS DE DESCARGA [CM]"
           
           If CARGARDATOS% > 0 Then
              MSF1.TextMatrix(1, 2) = TXTGRAMAJETAPA.TEXT
              MSF1.TextMatrix(2, 2) = TXT_VPDPOLLERAGM2.TEXT
              MSF1.TextMatrix(3, 2) = TXTGRAMAJECUERPO.TEXT
              MSF1.TextMatrix(4, 2) = TXTGRAMAJECUERPO.TEXT
              MSF1.TextMatrix(5, 2) = TXTGRAMAJEFONDO.TEXT
              MSF1.TextMatrix(6, 2) = TXTGRAMAJEFONDO.TEXT
              MSF1.TextMatrix(7, 2) = TXTVCLARGOCM.TEXT
              MSF1.TextMatrix(8, 2) = TXTVDLARGOCM.TEXT
              '
              'COMPLETO LA COLUMNA MODULO DE LA TABLA CONFIGBB
              SQLX$ = "SELECT TAPAPOLLERA,POLLERA,CUERPOLAT,FONDOPANEL,VALVCARGA,VALVDESCAR,LATERAL2,UPANEL "
              SQLX$ = SQLX$ + " FROM CONFIGBB WITH(NOLOCK)  WHERE CODBIGBAG = '" & TRIM$(TXTCODCONFECCION) & "'"
              Set RS = READSET(SQLX$)
              With RS
                If Not .EOF Then
                    MSF1.TextMatrix(1, 5) = SAFETEXT$(!TAPAPOLLERA)
                    MSF1.TextMatrix(2, 5) = SAFETEXT$(!Pollera)
                    MSF1.TextMatrix(3, 5) = SAFETEXT$(!CUERPOLAT)
                    MSF1.TextMatrix(4, 5) = SAFETEXT$(!LATERAL2)
                    MSF1.TextMatrix(5, 5) = SAFETEXT$(!FONDOPANEL)
                    MSF1.TextMatrix(6, 5) = SAFETEXT$(!UPANEL)
                    MSF1.TextMatrix(7, 5) = SAFETEXT$(!VALVCARGA)
                    MSF1.TextMatrix(8, 5) = SAFETEXT$(!VALVDESCAR)
                End If
              End With
              RS.Close: Set RS = Nothing
              '
'              If ANCHO = LARGO Then 'SOLO EN LOS CASOS QUE ANCHO SEA IGUAL A LARGO
                    'CARGAR DATOS DE CORTE - LEO LAS FORMULAS
                    SQLX$ = "SELECT TAPAPOLLERA,POLLERA,CUERPOLAT,FONDOPANEL,LATERAL2,UPANEL "
                    SQLX$ = SQLX$ + " FROM CALCUCORTE WITH(NOLOCK)  WHERE CODBIGBAG = '" & TRIM$(TXTCODCONFECCION) & "'"
                    Set RS = READSET(SQLX$)
                    With RS
                      If Not .EOF Then
                          formulaTapa$ = SAFETEXT$(!TAPAPOLLERA)
                          formulaPollera$ = SAFETEXT$(!Pollera)
                          formulaCuerpo$ = SAFETEXT$(!CUERPOLAT)
                          formulaLateral$ = SAFETEXT$(!LATERAL2)
                          formulaFondo$ = SAFETEXT$(!FONDOPANEL)
                          formulaUPanel$ = SAFETEXT$(!UPANEL)
                      End If
                    End With
                    RS.Close: Set RS = Nothing
                    '
                    'ASIGNO LOS RESULTADOS DE LAS FORMULAS A CADA FILA EN EL CAMPO CORTE DE LA GRILLA
                    MSF1.TextMatrix(1, 6) = CORTE_CONTEPLAST(formulaTapa$, ALTO, ANCHO, Largo, "TAPA", "F12.2")
                    MSF1.TextMatrix(2, 6) = CORTE_CONTEPLAST(formulaPollera$, ALTO, ANCHO, Largo, "POLLERA", "F12.2")
                    MSF1.TextMatrix(3, 6) = CORTE_CONTEPLAST(formulaCuerpo$, ALTO, ANCHO, Largo, "CUERPO", "F12.2")
                    MSF1.TextMatrix(4, 6) = CORTE_CONTEPLAST(formulaLateral$, ALTO, ANCHO, Largo, "LATERALES", "F12.2")
                    MSF1.TextMatrix(5, 6) = CORTE_CONTEPLAST(formulaFondo$, ALTO, ANCHO, Largo, "FONDO", "F12.2")
                    MSF1.TextMatrix(6, 6) = CORTE_CONTEPLAST(formulaUPanel$, ALTO, ANCHO, Largo, "UPANEL", "F12.2")
                    '
                    'COPIO EN LA ULTIMA COLUMNA NO VISIBLE PARA VALIDAR LUEGO SI SE CAMBIO ALGUNO
                    MSF1.TextMatrix(1, 10) = MSF1.TextMatrix(1, 6)
                    MSF1.TextMatrix(2, 10) = MSF1.TextMatrix(2, 6)
                    MSF1.TextMatrix(3, 10) = MSF1.TextMatrix(3, 6)
                    MSF1.TextMatrix(4, 10) = MSF1.TextMatrix(4, 6)
                    MSF1.TextMatrix(5, 10) = MSF1.TextMatrix(5, 6)
                    MSF1.TextMatrix(6, 10) = MSF1.TextMatrix(6, 6)

'              End If
              '
              
              
              
           End If
End Sub



Sub CARGAMSF2(CARGARDATOS%)


           Campos$ = "COMPONENTE/A20;FICHA TECNICA(g-m2)/A20;CODIGO/A16;DESCRIPCION/A24;MODULO/F12;CORTE/F10.2;CONSUMO (cm)/f12.2;PRECIO (u$s*m)/F12.2;PRECIO (u$s)/F12.2"
           Call CARGA_ENCABEZADO(msf2, Campos$, Me)
           msf2.Rows = 7
           msf2.TextMatrix(1, 1) = "CINTILLAS D [CM]"
           msf2.TextMatrix(2, 1) = "CINTILLA DS [CM]"
           msf2.TextMatrix(3, 1) = "CINTILLAS C [CM]"
           msf2.TextMatrix(4, 1) = "CINTILLA CS [CM]"
           msf2.TextMatrix(5, 1) = "CINTILLA POLLERA [CM]"
           msf2.TextMatrix(6, 1) = "CINTILLA LINER [CM]"
           
           If CARGARDATOS% > 0 Then
              msf2.TextMatrix(1, 6) = TXT_VPDVDESCARGACM_EDIT
              msf2.TextMatrix(2, 6) = TXT_VPDVDESCSEGCM_EDIT
              msf2.TextMatrix(3, 6) = TXT_VPDCargaCm_EDIT
              msf2.TextMatrix(4, 6) = TXT_VPDvcsegcm_EDIT
              msf2.TextMatrix(5, 6) = TXT_VPDPOLLERA_EDIT
              msf2.TextMatrix(6, 6) = TXTCINTILLALINER_EDIT
              '
              'COMPLETO LA COLUMNA MODULO DE LA TABLA CONFIGBB
              'CINTDESC, CINTDESCSEG,CINTCARGPOLL, CINTCARGSEG,  CINTLINER,
              SQLX$ = "SELECT CINTDESC, CINTDESCSEG,CINTCARGPOLL, CINTCARGSEG,  CINTLINER,CINTIPOLLERA "
              SQLX$ = SQLX$ + " FROM CONFIGBB WITH(NOLOCK)  WHERE CODBIGBAG = '" & TRIM$(TXTCODCONFECCION) & "'"
              Set RS = READSET(SQLX$)
              With RS
                If Not .EOF Then
                    msf2.TextMatrix(1, 5) = SAFETEXT$(!CINTDESC)
                    msf2.TextMatrix(2, 5) = SAFETEXT$(!CINTDESCSEG)
                    msf2.TextMatrix(3, 5) = SAFETEXT$(!CINTCARGPOLL)
                    msf2.TextMatrix(4, 5) = SAFETEXT$(!CINTCARGSEG)
                    msf2.TextMatrix(5, 5) = SAFETEXT$(!CINTIPOLLERA)
                    msf2.TextMatrix(6, 5) = SAFETEXT$(!CINTLINER)
                End If
              End With
              RS.Close: Set RS = Nothing
           
           End If
End Sub


Sub CARGAMSF3(CARGARDATOS%)


           Campos$ = "COMPONENTE/A20;FICHA TECNICA(g-m2)/A20;CODIGO/A16;DESCRIPCION/A24;MODULO/F12;CORTE/F10.2;CONSUMO (cm)/f12.2;PRECIO (u$s*m)/F12.2;PRECIO (u$s)/F12.2"
           Call CARGA_ENCABEZADO(MSF3, Campos$, Me)
           MSF3.Rows = 4
           MSF3.TextMatrix(1, 1) = "MAMPARO [CM]"
           MSF3.TextMatrix(2, 1) = "ESLINGAS [CM]"
           
           If CARGARDATOS% > 0 Then
              MSF3.TextMatrix(1, 2) = TXT_VPDTELAMAMPARO
              MSF3.TextMatrix(2, 2) = TXTESLINGAOJOCM
              '
              'COMPLETO LA COLUMNA MODULO DE LA TABLA CONFIGBB
              'CINTDESC, CINTDESCSEG,CINTCARGPOLL, CINTCARGSEG,  CINTLINER,
              SQLX$ = "SELECT MAMPARO,ESLINGA,ESLPERIM "
              SQLX$ = SQLX$ + " FROM CONFIGBB WITH(NOLOCK)  WHERE CODBIGBAG = '" & TRIM$(TXTCODCONFECCION) & "'"
              Set RS = READSET(SQLX$)
              With RS
                If Not .EOF Then
                    MSF3.TextMatrix(1, 5) = SAFETEXT$(!Mamparo)
                    MSF3.TextMatrix(2, 5) = SAFETEXT$(!Eslinga)
                End If
              End With
              RS.Close: Set RS = Nothing
              MSF3.TextMatrix(1, 6) = TXT_VPDLARGOMAMAPAROCM
                 ESLG# = XVALO(TXTESLINGAOJOCM) * 2 + XVALO(TXTESLINGAMEDIDA1CM) + XVALO(TXTESLINGAMED2EDITABLE)
              MSF3.TextMatrix(2, 6) = TRIM$(FORMATNUM$(ESLG#, "F12.2"))
              
           End If
End Sub

Sub CARGAMSF4(CARGARDATOS%)


           Campos$ = "COMPONENTE/A20;FICHA TECNICA(g-m2)/A20;CODIGO/A16;DESCRIPCION/A24;MODULO/F12;CORTE/F10.2;CONSUMO (cm)/f12.2;PRECIO (u$s*m)/F12.2;PRECIO (u$s)/F12.2"
           Call CARGA_ENCABEZADO(MSF4, Campos$, Me)
           MSF4.Rows = 4
           MSF4.TextMatrix(1, 1) = "HILO1 [m] BIG BAG"
           MSF4.TextMatrix(2, 1) = "HILO2 [m] BIG BAG"
           MSF4.TextMatrix(3, 1) = "HILO3 [m] BIG BAG"
           If CARGARDATOS% > 0 Then
              MSF4.TextMatrix(1, 2) = TXT_VPDTELAMAMPARO
              '
              'COMPLETO LA COLUMNA MODULO DE LA TABLA CONFIGBB
              'CINTDESC, CINTDESCSEG,CINTCARGPOLL, CINTCARGSEG,  CINTLINER,
              SQLX$ = "SELECT HILO1,HILO2,HILO3 "
              SQLX$ = SQLX$ + " FROM CONFIGBB WITH(NOLOCK)  WHERE CODBIGBAG = '" & TRIM$(TXTCODCONFECCION) & "'"
              Set RS = READSET(SQLX$)
              With RS
                If Not .EOF Then
                    MSF4.TextMatrix(1, 5) = SAFETEXT$(!HILO1)
                    MSF4.TextMatrix(2, 5) = SAFETEXT$(!HILO2)
                    MSF4.TextMatrix(3, 5) = SAFETEXT$(!HILO3)
                End If
              End With
              RS.Close: Set RS = Nothing
              
           End If
End Sub



Sub COPIAR_ESTRUCTURA(CI1%, CI1NUEVO%)
''        Dim ITEMX(), CODIDATOX(), INDIOBJX(), UsoBrutoX(), CORTEX(), OBSERVACIONX(), COMENTARIOX$()
''        ReDim Preserve ITEMX(1), CODIDATOX(1), INDIOBJX(1), UsoBrutoX(1), CORTEX(1), OBSERVACIONX(1), COMENTARIOX$(1)
''        IND& = 0
''        SQLX$ = "SELECT * FROM ESPECIFIC WITH(NOLOCK) where COD_INTE= " & CI1%
''        SQLX$ = SQLX$ + " ORDER BY INDIOBJ ASC"
''        Set RS = READSET(SQLX$)
''        HAYESTRUCTURA% = 0
''        With RS
''          Do While Not .EOF
''             IND& = IND& + 1
''             ReDim Preserve ITEMX(IND&), CODIDATOX(IND&), INDIOBJX(IND&), UsoBrutoX(IND&), CORTEX(IND&), OBSERVACIONX(IND&), COMENTARIOX$(IND&)
''             ITEMX(IND&) = SAFETEXT$(!Item)
''             CODIDATOX(IND&) = SAFETEXT$(!CODIDATO)
''             INDIOBJX(IND&) = SAFETEXT$(!INDIOBJ)
''             UsoBrutoX(IND&) = SAFETEXT$(!USOBRUTO)
''             CORTEX(IND&) = XVALO(!CORTE)
''             OBSERVACIONX(IND&) = SAFETEXT$(!Observacion)
''             COMENTARIOX$(IND&) = SAFETEXT$(!COMENTARIO)
''             HAYESTRUCTURA% = 1
''            .MoveNext
''          Loop
''        End With
''        RS.Close: Set RS = Nothing
''
''        If HAYESTRUCTURA% < 1 Then Exit Sub
''
''        COD$ = CODEXT$(CI1NUEVO%)
''
''        Dim RS1 As ADODB.Recordset
''        Set RS1 = New ADODB.Recordset
''        SQLX$ = "Select * from ESPECIFIC  where COD_INTE=" & CI1NUEVO%
''        RS1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
''        With RS1
''         For J& = 1 To UBound(ITEMX)
''
''           .AddNew
''           !ID = J&
''           !CodiEmpr = CodiEmp%
''           !cod_inte = CI1NUEVO%
''           !Codigo = COD$
''           !Descripcion = ""
''           If J& = 1 Then !Descripcion = Left$(Me.TXTDESC(J&), 64)
''           !Descripcion = DESCR$
''           !Item = ITEMX(J&)
''           !CODIDATO = CODIDATOX(J&)
''           !PREPARO = ""
''           !REVISO = ""
''           !APROBO = ""
''           !FEPREPARO = CVDAT(0)
''           !FEPREPARO = CVDAT(0)
''           !FEPREPARO = CVDAT(0)
''           !CODCLIENTE = ""
''           !MarBorra = 0
''           !INDIOBJ = INDIOBJX(J&)
''           !USOBRUTO = UsoBrutoX(J&)
''           !CORTE = CORTEX(J&)
''           If J& = 30 Then
''             !Observacion = TXTOBS(J&).TEXT
''            Else
''             !Observacion = ""
''             !COMENTARIO = COMENTARIOX$(J&)
''           End If
''
''           .Update
''          Next
''          End With
''          RS1.Close
''          Set RS1 = Nothing
          
'           CACAM% = RMaster.Fields.Count
'
'           For KU% = 1 To CACAM%
'             DATO$ = SAFETEXT$(RMaster.Fields(KU% - 1).Value)
'             NOCAMPO$ = TRIM$(UCase$(RMaster.Fields(KU% - 1).Name))
'
'             If NOCAMPO$ <> "CODIGO" Then
'                If NOCAMPO$ <> "DESCRIPCION" Then
'                   If NOCAMPO$ <> "CODINT" Then
'                      If NOCAMPO$ <> "ALIAS" Then
'                         If NOCAMPO$ <> "UMEDIDA" Then
'                           CONDI$ = "CODINT=" & CI_NEW%
'                           Call ACTUREGISTRO("MASTER", NOCAMPO$, CONDI$, DATO$, REGAF&, "NOMESS")
'                         End If
'                      End If
'                   End If
'                End If
'              End If
'           Next KU%
'        End If
'       End If
'
'    SQLX$ = "INSERT INTO ESPECIFIC(ID"
'    SQLX$ = SQLX$ + ",1" & CodiEmp%
'    SQLX$ = SQLX$ + "," & CI1NUEVO%
'    SQLX$ = SQLX$ + ",'" & CODEXT$(CI1NUEVO%) & "'"
'    SQLX$ = SQLX$ + ",'" & DESCRIT0$(CI1NUEVO%) & "'"
'    SQLX$ = SQLX$ + ", Item, CODIDATO, '','','' "
'    SQLX$ = SQLX$ + ",'" & CVDAT(0) & "'"
'    SQLX$ = SQLX$ + " ,'" & CVDAT(0) & "'"
'    SQLX$ = SQLX$ + ",'" & CVDAT(0) & "'"
'    SQLX$ = SQLX$ + ", COMENTARIO, CODCLIENTE , OBSERVACION, MARBORRA, INDIOBJ, UsoBruto, CORTE)"
'    SQLX$ = SQLX$ + " SELECT ID, CodiEmpr, COD_INTE, Codigo, Descripcion, ITEM, CODIDATO, PREPARO, REVISO, APROBO, FEPREPARO, FEREVISO, FEAPROBO, COMENTARIO,"
'    SQLX$ = SQLX$ + "CODCLIENTE , OBSERVACION, MARBORRA, INDIOBJ, UsoBruto, CORTE FROM ESPECIFIC WHERE COD_INTE = " & CI1%
'    FLEXCONN.Execute SQLX$
End Sub


Function FORMATOCODIGOCONTEPLAST$(NC, CI1%, ESNUEVAVERSION%)
   FORMATOCODIGOCONTEPLAST$ = ""
   If NC < 1 And CI1% = 0 Then Exit Function
   '123456-89-11
   NC1& = NC
   PARTENROCLIENTE$ = FORMATOCON0$(NC1&, 6)
   If ESNUEVAVERSION% < 1 Then
      PARTEORDEN$ = FORMATOCON0$(NUEVOORDENCODIGOMASTERDELCLIENTE%(NC1&), 2)
   Else
      PARTEORDEN$ = Mid$(CODEXT$(CI1%), 8, 2)
      If NC1& <> XVALO(PARTENROCLIENTE$) Then
          PARTEORDEN$ = FORMATOCON0$(NUEVOORDENCODIGOMASTERDELCLIENTE%(NC1&), 2)
      End If
   End If
   PARCIALCODIGO$ = PARTENROCLIENTE$ & "-" & PARTEORDEN$ & "-"
   If CI1% = 0 Then 'CASO QUE NO TENGA CODIGO DEL CUAL GENERAR NUEVA VERSION
      FORMATOCODIGOCONTEPLAST$ = PARTENROCLIENTE$ & "-" & PARTEORDEN$ & "-00"
   Else
      'CASO QUE TENGA CODIGO Y N.CLIENTE
      PARTEVERSION$ = FORMATOCON0$(GENERAR_NUEVA_VERSION_NUMERICA%(PARCIALCODIGO$), 2)
      FORMATOCODIGOCONTEPLAST$ = PARTENROCLIENTE$ & "-" & PARTEORDEN$ & "-" & PARTEVERSION$
   End If
End Function
Function CREARCODIGO99000$(NC)
  CREARCODIGO99000$ = ""
  NC1& = NC
  NUEVOINDICE& = GENERAR_NUEVA_INDICE_CTA_99000%(NC) ' 10
  INDICETEX$ = FORMATOCON0$(NUEVOINDICE&, 4) '0010
  INDICETEX$ = Left$(INDICETEX$, 2) & "-" & Right(INDICETEX$, 2) '00-10
  CREARCODIGO99000$ = FORMATOCON0$(NC1&, 6) & "-" & INDICETEX$ '099000-00-10
End Function

Sub GUARDARARCHADJUNTOS(FTEC&)
   'MARCA DE BORRADO
   Call ACTUREGISTRO("DOCURELFT", "MARBORRA", "NUMRECETACTP = " & FTEC&, 1, REG&)
   '
   'RECORRO EL LIST PARA GRABAR
   For i& = 1 To Combo2.ListCount
      Z$ = Combo2.List(i& - 1)
      DESCRI$ = UCase$(TRIM$(Mid$(Z, 1, 24)))
      RUTAX$ = TRIM$(Mid$(Z, 32, 512))
      SQLX$ = "INSERT INTO  DOCURELFT (NUMRECETACTP,CODINT ,Descripcion, Ruta, Marborra)"
      SQLX$ = SQLX$ + "VALUES ( " & FTEC&
      SQLX$ = SQLX$ + "," & CODINT(TXTCODIGO) 'CI%
      SQLX$ = SQLX$ + ",'" & DESCRI$ & "'"
      SQLX$ = SQLX$ + ",'" & TRIM$(RUTAX$) & "'"
      SQLX$ = SQLX$ + ",0"
      SQLX$ = SQLX$ + ")"
      FLEXCONN.Execute (SQLX$)
   Next i&
   '
   'BORRO LOS REGISTROS MARCADOS CON 1
   Call BORRAREGISTROS("DOCURELFT", "MARBORRA = 1 AND NUMRECETACTP = " & FTEC&, REGAF&, "NOMESS")
   


End Sub

Sub LIMPIACASILLEROS()

  Me.TXTPREPARO = ""
  Me.TXTAPROBO = "": Me.TXTFECHAPROBO = ""
  
  List1.Clear
End Sub









Function DIFIERECORTE&(MSF As msFlexGrid, COLEDITABLE%, COLRESPALDO%)
    'RECORRE LA GRILLA, COMPARA EL DATO DE LA COLUMNA EDITABLE(6) CON EL DE LA COLUMNA DEL VALOR ORIGINAL(10)
    'SI ENCUENTRA DIFERENCIA DEVUELVE EL NUMERO DE FILA DONDE DIFIERE UNO DEL OTRO
    DIFIERECORTE& = 0
    For REG& = 1 To MSF.Rows - 1
       DATORESPALDO = FORMATNUM$(XVALO(MSF.TextMatrix(REG&, COLRESPALDO%)), "F10.2")
       DATOEDITADO = FORMATNUM$(XVALO(MSF.TextMatrix(REG&, COLEDITABLE%)), "F10.2")
       If XVALO(DATORESPALDO) > 0 Then
          If DATORESPALDO <> DATOEDITADO Then
             DIFIERECORTE& = REG&
             Exit Function
          End If
       End If
    Next REG&
End Function

Function NROSOLICITUDCOTIZACION(NROCOT&)
   'DEVUELVE EL NRO DE SOLICITUD DE COTIZACION DE TABLA SOLICOTI SEGUN EL NRO DE COTIZACION DE TABLA COTIZACTP
   CONDI$ = "NROCOTIZACION= " & NROCOT&
   NROSOLICITUDCOTIZACION = XVALO(VALOBADA("SOLICOTI", "SOLICOTI_ID", CONDI$, "NOMESS"))
End Function

Function VALIDASISEMODIFICO%()
    If XVALO(Text5) > 0 Then
     If MODIFIC% = 1 Then
          If COMALTER%("Se han Modificado Datos de La F.Técnica Activa\\Registrar Cambios\Continuar Sin Grabar") = 1 Then
            Call MnuGuardar_Click
          End If
          MODIFIC% = 0
     End If
    End If
   VALIDASISEMODIFICO% = 0
   MODIFIC% = 0
End Function



Sub VERIFICARWINRAR()
On Error Resume Next
Dim MiObjeto As Object, ret As String
Set MiObjeto = CreateObject("Wscript.Shell")
ret = MiObjeto.RegRead("HKEY_CLASSES_ROOT\WinRAR\shell\open\command\")
WinRar = Left(ret, InStrRev(ret, " "))

ret = MiObjeto.RegRead("HKEY_CLASSES_ROOT\WinZip\shell\open\command\")
WinZip = Left(ret, InStrRev(ret, " "))
Set MiObjeto = Nothing
If WinRar & WinZip = "" Then
 MsgBox "No se encuentra instalado ni WinZip ni WinRar :("
 End
 On Error GoTo 0
End If
End Sub

Private Sub CMD1_Click()
'   If FIRMAELECTRONICAOK% <= 0 Then Exit Sub

   X = LISTASELECGENERAL("TABLASI_NO", , , 1)
   TXTJUSTA = TRIM$(RESP_ZELE_VECT(2))
End Sub

Private Sub CMDFECHA_Click()
   Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      TXTFECHA.TEXT = VALACT1$("SELFECHA")
   End If

End Sub

Private Sub Command1_Click()

   USNU% = USNBR% Mod 100
   VENDESEGUNUSUARIO% = XVALO(Control("LOGVEND", SAFETEXT(USNU%)))
   '
   NC = NCLLICONTEPLAST(VENDESEGUNUSUARIO%)
   If NC < 1 Then Exit Sub
   Text1.TEXT = TRIM$(Str$(NC))
End Sub

Private Sub Command10_Click()
   X = LISTASELECGENERAL("CTP.GRAMAJE", , , 1)
   Me.TXTGRAMAJEFONDO = TRIM$(RESP_ZELE_VECT(2))
   Call VALORESPORDEFECTO

End Sub

Private Sub Command11_Click()
   X = LISTASELECGENERAL("CTP.TAPA_GR", , , 1)
   Me.TXTGRAMAJETAPA = TRIM$(RESP_ZELE_VECT(2))
   Call VALORESPORDEFECTO

End Sub


Private Sub Command12_Click()
   X = LISTASELECGENERAL("TABLASI_NO", , , 1)
   Me.TXTLAMFONDOLAMINADO = TRIM$(RESP_ZELE_VECT(2))
   Call VALORESPORDEFECTO
   
End Sub

Private Sub Command13_Click()
   X = LISTASELECGENERAL("TABLASI_NO", , , 1)
   Me.TXTLAMCUERPOLAMINADO = TRIM$(RESP_ZELE_VECT(2))
   Call VALORESPORDEFECTO

End Sub


Private Sub Command14_Click()

   X = LISTASELECGENERAL("CTP.PDCOTIPO", , , 1)
   Me.TXTPDOCPORTADOCUMENTOTIPO = TRIM$(RESP_ZELE_VECT(2))
   Call VALORESPORDEFECTO

End Sub


Private Sub Command15_Click()
   X = LISTASELECGENERAL("CTP.PD", , , 1)
   Me.TXTPDOCPDOCUMENTOUBIC = TRIM$(RESP_ZELE_VECT(2))
   Call VALORESPORDEFECTO

End Sub


Private Sub Command16_Click()
   X = LISTASELECGENERAL("TABLASI_NO", , , 1)
   Me.TXTCOMPARTCOMPARTIMENTADO = TRIM$(RESP_ZELE_VECT(2))
   Call VALORESPORDEFECTO

End Sub

Private Sub Command17_Click()
   X = LISTASELECGENERAL("TABLASI_NO", , , 1)
   Me.TXTLINERLINER = TRIM$(RESP_ZELE_VECT(2))
   Call VALORESPORDEFECTO

End Sub

Private Sub Command18_Click()
   X = LISTASELECGENERAL("CTP.LINT", , , 1)
   Me.TXTLINERTIPO = TRIM$(RESP_ZELE_VECT(2))
   Call VALORESPORDEFECTO

End Sub


Private Sub Command19_Click()
   X = LISTASELECGENERAL("CTP.EMPRESA", , , 1)
   Me.TXTDETALLEETIQUETA = TRIM$(RESP_ZELE_VECT(2))
   Call VALORESPORDEFECTO

End Sub


Private Sub Command2_Click()
    Call SELECHO("VENDEDOR")
    VDEVU$ = TRIM$(VALACT1$("VENDEDOR"))
    If VDEVU$ <> "" Then
       Text4.TEXT = VDEVU$
    End If

End Sub

Private Sub Command20_Click()
   X = LISTASELECGENERAL("CTP.FAR", , , 1)
   Me.TXTEMBALAJEFARDO = TRIM$(RESP_ZELE_VECT(2))
   Call VALORESPORDEFECTO

End Sub

Private Sub Command21_Click()
   X = LISTASELECGENERAL("CTP.COSTURA", , , 1)
   Me.TXTCOSTURAEXTERNA = TRIM$(RESP_ZELE_VECT(2))

End Sub

Private Sub Command22_Click()
   X = LISTASELECGENERAL("CTP.PAB", , , 1)
   Me.TXTCOSTURACONPABILO = TRIM$(RESP_ZELE_VECT(2))
End Sub

Private Sub Command23_Click()
   X = LISTASELECGENERAL("TABLASI_NO", , , 1)
   Me.TXTESLINGAPERIMETRAL = TRIM$(RESP_ZELE_VECT(2))

End Sub

Private Sub Command24_Click()
   X = LISTASELECGENERAL("CTP.ES", , , 1)
   Me.TXTESLINGAANCHO = TRIM$(RESP_ZELE_VECT(2))


End Sub

Private Sub Command25_Click()
   X = LISTASELECGENERAL("CTP.ESCOST", , , 1)
   Me.TXTTIPOCOSTURA = TRIM$(RESP_ZELE_VECT(2))
   Call VALORESPORDEFECTO

End Sub

Private Sub Command26_Click()
   X = LISTASELECGENERAL("CTP.VALV", , , 1)
   Me.TXTVALVVDDIAMETROCM = TRIM$(RESP_ZELE_VECT(2))
   Call VALORESPORDEFECTO

End Sub

Private Sub Command27_Click()
   X = LISTASELECGENERAL("TABLASI_NO", , , 1)
   TXTPOLLERALAMINADA = TRIM$(RESP_ZELE_VECT(2))
   Call VALORESPORDEFECTO

End Sub

Private Sub Command28_Click()
   X = LISTASELECGENERAL("TABLASI_NO", , , 1)
   Me.TXTTELALAMINADA = TRIM$(RESP_ZELE_VECT(2))
   Call VALORESPORDEFECTO

End Sub

Private Sub Command29_Click()
   X = LISTASELECGENERAL("CTP.LI ", , , 1)
   Me.TXTCINTILLALINER = TRIM$(RESP_ZELE_VECT(2))
   Call VALORESPORDEFECTO

End Sub

Private Sub Command3_Click()
   Erase RESP_ZELE_VECT
   X = LISTASELECGENERAL("CODBB", "A12", "A128")
   If X <> "" Then
     Me.TXTCODCONFECCION = TRIM$(RESP_ZELE_VECT(1))
     Me.TXTDESCRICODCONFECCION = TRIM$(RESP_ZELE_VECT(2))
   End If
End Sub

Private Sub Command30_Click()
   X = LISTASELECGENERAL("CTP.LINER", , , 1)
   Me.TXTMICRONAJELINER = TRIM$(RESP_ZELE_VECT(2))
   Call VALORESPORDEFECTO


End Sub

Private Sub Command31_Click()
   X = LISTASELECGENERAL("CTP.ETIQ", , 11000, 1)
   Me.TXTDETAETIQUETA2 = TRIM$(RESP_ZELE_VECT(2))
    Call VALORESPORDEFECTO

End Sub

Private Sub Command32_Click()
   X = LISTASELECGENERAL("CTP.PALLET", , , 1)
   Me.TXTPALLET = TRIM$(RESP_ZELE_VECT(2))
   Call VALORESPORDEFECTO

End Sub

Private Sub Command33_Click()
   X = LISTASELECGENERAL("CTP.EMB", , , 1)
   Me.TXTUNPALLET = TRIM$(RESP_ZELE_VECT(2))

End Sub

Private Sub Command34_Click()

   X = LISTASELECGENERAL("CTP.PDTA", , , 1)

   Me.TXTDOCUMENTOUNIDAD = TRIM$(RESP_ZELE_VECT(2))
   Call VALORESPORDEFECTO

End Sub

Private Sub Command35_Click()
   X = LISTASELECGENERAL("CTP.CARAS", , , 1)

   Me.TXTPORTADOCUMEDIDA = TRIM$(RESP_ZELE_VECT(2))
   Call VALORESPORDEFECTO

End Sub

Private Sub Command36_Click()
   X = LISTASELECGENERAL("CTP.VALV", , , 1)
   If XVALO(RESP_ZELE_VECT(3)) > 32767 Then
       Call COMUNI("Atención se ha Detectado un Valor Ingresado en Esta Tabla Fuera de Rango\El Campo Parametro Desc Seg (cm) Soporta Nros. Enteros Hasta 32767\Por Favor Corrija el Mismo")
   End If
   Me.TXTVALVVCDIAMETROCM = TRIM$(RESP_ZELE_VECT(2))
   
   Call VALORESPORDEFECTO
End Sub

Private Sub Command37_Click()
   X = LISTASELECGENERAL("TABLASI_NO", , , 1)
   TXTVALVVCLARGOCM = TRIM$(RESP_ZELE_VECT(2))
    Call VALORESPORDEFECTO

End Sub

Private Sub Command38_Click()
   X = LISTASELECGENERAL("CTP.POLLERA", , , 1)
   TXTPOLLERALARGO = TRIM$(RESP_ZELE_VECT(2))
   Call VALORESPORDEFECTO

End Sub



Private Sub Command39_Click()
   Command39.Enabled = False
   NC = XVALO(Text1)
   If TRIM$(TXTCODIGO) = "" Then
     Call COMUNI("Debe Haber un Código Activo")
     Call PINTATEXT(TXTCODIGO)
     GoTo 99
   End If
   '
   FTEC& = XVALO(Text5)
   If FTEC& < 1 Then
     Call COMUNI("Debe Haber una Ficha Técnica Activa")
     Call PINTATEXT(Text5)
     GoTo 99
   End If

   If NC < 1 Then
     Call COMUNI("Debe Haber un Cliente Activo")
     Call PINTATEXT(Text1)
     GoTo 99
   End If
   '
   NUVENDX = XVALO(Text4)
   If NUVENDX < 1 Then
     Call COMUNI("Debe Haber un Vendedor Activo")
     Call PINTATEXT(Text4)
     GoTo 99
   End If
   
   CODBB = TRIM$(TXTCODCONFECCION)
   If CODBB = "" Then
     Call COMUNI("Debe Haber un Código de Confección Activo")
     Call PINTATEXT(TXTCODCONFECCION)
     GoTo 99
   End If
   '
   If CantRegistros&("DOCURELFT", "NUMRECETACTP = " & FTEC&, "NOMESS") < 1 Then
      OPX% = COMALTER%("Atención No se Han Ingresado Imágenes a la Ficha Técnica Activa\\Salir y Aplicar\Continuar")
      If OPX% < 2 Then GoTo 99
   End If
   
   CANTI = XVALO(Text2)
   
20 CONFIRMACANTI$ = TRIM$(InputBox$("Confirme Cantidad ", "Cantidad", Text2))
   If CONFIRMACANTI$ = "" Then Exit Sub
   If XVALO(CONFIRMACANTI$) < 1 Then
      Call COMUNI("Cantidad Ingresada No Válida")
   End If
   
   If TRIM$(TXTRAZOCLI) = "" Then
      If NC <> 99000 Then
         TXTRAZOCLI = rasocli$(NC)
      End If
   End If

    On Error GoTo ERROR_GUARDAR
    FLEXCONN.BeginTrans

   
    SQLX$ = "INSERT INTO  SOLICOTI (NUVEND ,CANTIDAD, CODBIGBAG, CODIMASTER, COTIZAR, NUMRECETACTP, NROCOTIZACION, CODINT "
    SQLX$ = SQLX$ + ", NUCLIEN, CodiEmpr, MARBORRA, FECHEMISI,RASO_CLI,DESCRIMASTER)"
    SQLX$ = SQLX$ + "VALUES ( " & NUVENDX 'NUVEND
    SQLX$ = SQLX$ + "," & XVALO(CONFIRMACANTI$) 'CANTIDAD
    SQLX$ = SQLX$ + ",'" & CODBB & "'" 'CODBIGBAG
    SQLX$ = SQLX$ + ",'" & TRIM$(TXTCODIGO) & "'" 'CODIMASTER
    SQLX$ = SQLX$ + ",1" 'COTIZAR
    SQLX$ = SQLX$ + " ," & XVALO(Text5) 'NUMRECETACTP
    SQLX$ = SQLX$ + " ,0" 'NROCOTIZACION
    SQLX$ = SQLX$ + " ," & CODINT%(TXTCODIGO) 'CODINT
    SQLX$ = SQLX$ + " ," & NC 'NUCLIEN
    SQLX$ = SQLX$ + " ," & CodiEmp%
    SQLX$ = SQLX$ + " ,0" 'MARBORRA
    SQLX$ = SQLX$ + " ,'" & FETEXCOR1$(HOY(HO$)) & "'" 'FECHEMISI
    SQLX$ = SQLX$ + " ,'" & TRIM$(Left$(TXTRAZOCLI, 64)) & "'" 'RASO_CLI
    SQLX$ = SQLX$ + " ,'" & TRIM$(Left$(TXTDESCRICOD, 64)) & "'" 'FECHEMISI
    SQLX$ = SQLX$ + ")"
    FLEXCONN.Execute (SQLX$)
    FLEXCONN.CommitTrans
    Call COMUNI("Grabación Realizado Exitosamente")
    GoTo 99
ERROR_GUARDAR:
        If Err <> 0 Then
           FLEXCONN.RollbackTrans
           MsgBox Err.Description
           On Error GoTo 0
           Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION")
           Exit Sub
        End If
   
99 Command39.Enabled = True
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   NC = XVALO(Text1)
   '
   If NC < 1 Then
      Call COMUNI("Para Solicitar Cotización Debe Tener Asignado un Cliente Destino")
      GoTo 99
   End If
   '
   If CantRegistros&("RECETACTP", "NUMRECETACTP = " & XVALO(Text5)) < 1 Then
      Call COMUNI("Debe Seleccionar Una Cotización Que esté Correctamente Guardada en al Base de Datos")
      GoTo 99
   End If
   
   CONDI$ = "NUMRECETACTP = " & XVALO(Text5)
   Call ACTUREGISTRO("RECETACTP", "STATUS", CONDI$, 1, REGAF&, "NOMESS")
   If REGAF& > 0 Then
      Call COMUNI("Se ha Rregistrado Como Pendiente de Aprobación")
   End If
   
99 Command4.Enabled = True
End Sub

Private Sub Command40_Click()
  Command40.Enabled = False
  NRO = SELEC_RECETA(0, 2, , -1)
  If NRO > 0 Then
     Text5 = NRO
  End If
  
99   Command40.Enabled = True
End Sub


Private Sub Command41_Click()
    Command41.Enabled = False
    CI1% = CODINT%(TRIM$(TXTCODIGO))
    If CI1% < 1 Then
       Call COMUNI("Código Inexistente o no Válido")
       GoTo 99
    End If
    
    CUERPO$ = Text6
    Call ANOTAHISMASTER(CI1%, CUERPO$)
    Call CARGAHISTORIAL(CI1%)
    
99  Command41.Enabled = True

End Sub

Private Sub Command42_Click()
   Command42.Enabled = False
   TXTAPROBO_DblClick
99 Command42.Enabled = True
End Sub
Private Sub Command43_Click()
    Command43.Enabled = False
    FTEC& = XVALO(Text5)
    If FTEC& < 1 Then
          Call COMUNI("Debe Haber una Ficha Técnica Activa")
          Call PINTATEXT(Text5)
          GoTo 99
    End If
      
    If CantRegistros&("RECETACTP", "NUMRECETACTP = " & XVALO(Text5)) < 1 Then
        Call COMUNI("Debe Seleccionar Una Ficha Técnica Que esté Correctamente Guardada en al Base de Datos")
        GoTo 99
    End If
    If Combo2.ListIndex >= 0 Then
      OPX% = COMALTER%("Confirmación de Eliminación Archivo Adjunto Seleccionado\\Cancelar\Si Eliminar")
      If OPX% = 2 Then
       Combo2.RemoveItem (Combo2.ListIndex)
       Call GUARDARARCHADJUNTOS(FTEC&)
      End If
    Else
       Call COMUNI("Atención No Hay Ningún Archivo Seleccionado")
       GoTo 99
    End If
    
99  Command43.Enabled = True
End Sub

Private Sub Command44_Click()
     COD$ = ListaSeleccionArticulos("", ActualizaListaSeleccion)
     ActualizaListaSeleccion = False
     If COD$ <> "" Then
       TXTCODIGO = COD$
     End If

End Sub




Private Sub Command45_Click()
   Command45.Enabled = False

   If Combo2.ListIndex >= 0 Then
     Linea$ = Combo2.List(Combo2.ListIndex)
     Archi$ = TRIM$(Mid$(Linea$, 25))
     If (EXISTE%(Archi$)) <> 1 Then
       Call COMUNI("Archivo Inexistente")
       Exit Sub
     Else
        Call ShellExecute(Me.hwnd, "Open", Archi$, "", "", 1)
        GoTo 99
     End If
   Else
      Call COMUNI("Atención No Hay Ningún Archivo Seleccionado")
      GoTo 99
   End If
    
99 Command45.Enabled = True
End Sub


Private Sub Command46_Click()
        Command46.Enabled = False
        ESNUEVO% = 0
        If lblNuevaCotizacion.Visible = True Then
          ESNUEVO% = 1
        End If
        '
        If XVALO(TXTNUMCOTIZACION) < 1 Then
           Call COMUNI("Debe Seleccionar Una Cotización")
           GoTo 99
        End If
        SQLX$ = "SELECT * FROM COTIZACTP "
        If ESNUEVO% < 1 Then
         SQLX$ = SQLX$ + " WHERE NROCOTIZACION  = " & XVALO(TXTNUMCOTIZACION)
        Else
         SQLX$ = SQLX$ + " WHERE NROCOTIZACION  < 1 "
         
        End If
        CODINCONJ% = CODINT%(TXTCODIGO)
        On Error GoTo ERROR_GUARDAR
        FLEXCONN.BeginTrans
        'GUARDA EN TABLA CABEZERA COTIZACTP
        Dim RS As ADODB.Recordset
        Set RS = New ADODB.Recordset
        RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        With RS
         If ESNUEVO% > 0 Then
             NEWCOTIZ = NumNuevaCotiza
             .AddNew
             !NROCOTIZACION = NEWCOTIZ
             !FECHEMISI = CVDAT(HOY(HO$))
             Else
               NEWCOTIZ = XVALO(TXTNUMCOTIZACION)
         End If
         !rasocli = TXTRAZOCLI
         !Status = 0
         !NUMRECETACTP = XVALO(Text5)
         !NUCLIEN = XVALO(Text1)
         !CODIMASTER = TXTCODIGO
         .Update
        End With
        RS.Close
        Set RS = Nothing
        
        'GUARDA EN TABLA COTI1 EL DETALLE DE LOS MSF
        Dim MSF As msFlexGrid
        Set MSF = MSF1
        Dim RS1 As ADODB.Recordset
        Set RS1 = New ADODB.Recordset
        
        SQLX$ = "SELECT * FROM COTI1"
        SQLX$ = SQLX$ + " WHERE NROCOTIZACION  < 1 "
        If ESNUEVO% < 1 Then
           Call ACTUREGISTRO("COTI1", "MARBORRA", "NROCOTIZACION = " & NEWCOTIZ, 1, REGAF&, "NOMES")
        End If
        
        RS1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        With RS1
            'RECORRE LOS 5 MSF Y GUARDA LOS DATOS
            'GUARDA EL NOMBRE DEL MSF Y EL ORDEN PARA PODER AGREGAR CUANDO SE LEVANTE LA INFO PARA MOSTRAR
            For J% = 1 To 5
              For REG& = 1 To MSF.Rows - 1
                 .AddNew
                 !NROCOTIZACION = NEWCOTIZ
                 !FECHEMISI = CVDAT(HOY(HO$))
                 !CODIEMPR = CodiEmp%
                 !COMPONENTE = SAFETEXT$(MSF.TextMatrix(REG&, 1))
                 !F_TECNICA = SAFETEXT$(MSF.TextMatrix(REG&, 2))
                 !Codigo = SAFETEXT$(MSF.TextMatrix(REG&, 3))
                 !Descripcion = SAFETEXT$(MSF.TextMatrix(REG&, 4))
                 !moduxbb = XVALO(MSF.TextMatrix(REG&, 5))
                 !corte = XVALO(MSF.TextMatrix(REG&, 6))
                 
                 !PRECIOXM = ROUND(MSF.TextMatrix(REG&, 8), 4)
                 !FECHEMISI = CVDAT(HOY(HO$))
                 !NAMEMSF = MSF.Name
                 !MARBORRA = 0
                 !ORDEN = REG&
                 !CODINTELEM = CODINT%(MSF.TextMatrix(REG&, 3))
                 !CODINTCONJ = CODINCONJ%
                 !Status = 0
                 If MSF.Name = "MSF1" Then !RESPACORTE = XVALO(MSF.TextMatrix(REG&, 10))
                   .Update
              Next REG&
              Select Case J%
                 Case 1
                   Set MSF = msf2
                 Case 2
                   Set MSF = MSF3
                 Case 3
                   Set MSF = MSF4
                 Case 4
                   Set MSF = MSF5
                 End Select
            Next J%
        End With
        RS1.Close
        Set RS1 = Nothing
        '
        'GRABO EL NRO DE SOLICITUD EN LA TABLA DE SOLICITUDES DE COTIZACION.
        NROSOLICIT& = XVALO(Me.LBLNROSOLICITUDCOTIZACION)
        If NROSOLICIT& > 0 Then
           SQLX$ = "UPDATE SOLICOTI SET NROCOTIZACION = " & NEWCOTIZ
           SQLX$ = SQLX$ + " WHERE SOLICOTI_ID = " & NROSOLICIT&
           FLEXCONN.Execute (SQLX$)
        End If
        
ERROR_GUARDAR:
        If Err <> 0 Then
           FLEXCONN.RollbackTrans
           MsgBox Err.Description
           On Error GoTo 0
           Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION")
           Call BORRAREGISTROS("COTI1", "NROCOTIZACION = " & NEWCOTIZ & " AND MARBORRA = 1", REGAF&, "NOMES")
           Exit Sub
        Else
           Call BORRAREGISTROS("COTI1", "NROCOTIZACION = " & NEWCOTIZ & " AND MARBORRA = 1", REGAF&, "NOMES")
           
           FLEXCONN.CommitTrans
           Call COMUNI("Grabación Realizado Exitosamente")
           lblNuevaCotizacion.Visible = False
           If XVALO(NEWCOTIZ) <> XVALO(TXTNUMCOTIZACION) Then TXTNUMCOTIZACION = XVALO(NEWCOTIZ)
           

           MODIFICCOT1% = 0
        End If
99      Command46.Enabled = True
End Sub




Private Sub Command47_Click()
    Command47.Enabled = False
    FTEC& = XVALO(Text5)
    If FTEC& < 1 Then
          Call COMUNI("Debe Haber una Ficha Técnica Activa")
          Call PINTATEXT(Text5)
          GoTo 99
    End If
      
    If CantRegistros&("RECETACTP", "NUMRECETACTP = " & XVALO(Text5)) < 1 Then
        Call COMUNI("Debe Seleccionar Una Ficha Técnica Que esté Correctamente Guardada en al Base de Datos")
        GoTo 99
    End If

    'BOTON ADD
    CARPETAX$ = Control("CARPIMAG", "UBICIMAG")
    If CARPETAX$ = "" Then
       Call COMUNI("Imposible Continuar\Falta Seleccionar Carpeta de Imagenes")
       GoTo 99
     End If
     
    On Error Resume Next
    X = GetAttr(CARPETAX$)
    If Err Then
        Call COMUNI("Imposible Continuar\Falta Seleccionar Carpeta de Imagenes o Ruta Inexistente")
        GoTo 99
    End If
    '
    On Error GoTo 0
   
    Call AGREGARRUTAIMAG(FTEC&, APRO%)
    If APRO% < 0 Then GoTo 99
    Call GUARDARARCHADJUNTOS(FTEC&)
'     'BOTON QUITAR
'     Case 1
'       If List2.ListIndex >= 0 Then
'          List2.RemoveItem (List2.ListIndex)
'       End If
'   End Select
   '
   '
   
   
99 Command47.Enabled = True

End Sub

Private Sub Command48_Click()
    Exit Sub
    
    NROCOT& = XVALO(TXTNUMCOTIZACION)
    
    NROFTEC& = XVALO(Text5)
    FILTX$ = TRIM$("102030-01") & ";" & TRIM$(Str$(NROFTEC&)) & ";" & TRIM$(Str$(NROCOT&))
    Call STDFORM("OCOR-CTP", FILTX$)
    
'    if !IMPRESION = si
    FILTX$ = TRIM$("102030-01") & ";" & TRIM$(Str$(NROFTEC&)) & ";" & TRIM$(Str$(NROCOT&))
    Call STDFORM("OIMP-CTP", FILTX$)
    'CUERPO
    CODICOTIZADO% = CodigodeCotizacion&(NROCOT&, "CUERPO")
    CRANGOEXT$ = CODEXT$(CODICOTIZADO%)
    If TRIM$(CRANGOEXT$) = "" Then CRANGOEXT$ = "-----"
    FILTX$ = TRIM$("102030-01") & ";" & TRIM$(Str$(NROFTEC&)) & ";" & TRIM$(Str$(NROCOT&)) & ";" & TRIM$(CRANGOEXT$)
    Call STDFORM("OCCU-CTP", FILTX$)
    
    'ESLINGA
    CODICOTIZADO = CodigodeCotizacion&(NROCOT&, "ESLINGAS [CM]")
    CRANGOEXT$ = CODEXT$(CODICOTIZADO)
    If TRIM$(CRANGOEXT$) = "" Then CRANGOEXT$ = "-----"
    FILTX$ = TRIM$("102030-01") & ";" & TRIM$(Str$(NROFTEC&)) & ";" & TRIM$(Str$(NROCOT&)) & ";" & TRIM$(CRANGOEXT$)
    Call STDFORM("OCES-CTP", FILTX$)
    
   'TAPAS
    CODICOTIZADO = CodigodeCotizacion&(NROCOT&, "TAPA")
    CRANGOEXT$ = CODEXT$(CODICOTIZADO)
    If TRIM$(CRANGOEXT$) = "" Then CRANGOEXT$ = "-----"
    FILTX$ = TRIM$("102030-01") & ";" & TRIM$(Str$(NROFTEC&)) & ";" & TRIM$(Str$(NROCOT&)) & ";" & TRIM$(CRANGOEXT$)
    Call STDFORM("OCTP-CTP", FILTX$)
  
  'FONDO
    CODICOTIZADO = CodigodeCotizacion&(NROCOT&, "FONDO")
    CRANGOEXT$ = CODEXT$(CODICOTIZADO)
    If TRIM$(CRANGOEXT$) = "" Then CRANGOEXT$ = "-----"
    FILTX$ = TRIM$("102030-01") & ";" & TRIM$(Str$(NROFTEC&)) & ";" & TRIM$(Str$(NROCOT&)) & ";" & TRIM$(CRANGOEXT$)
    Call STDFORM("OFDO-CTP", FILTX$)
  
  'MAMPAROS
    CODICOTIZADO = CodigodeCotizacion&(NROCOT&, "MAMPARO [CM]")
    CRANGOEXT$ = CODEXT$(CODICOTIZADO)
    If TRIM$(CRANGOEXT$) = "" Then CRANGOEXT$ = "-----"
    FILTX$ = TRIM$("102030-01") & ";" & TRIM$(Str$(NROFTEC&)) & ";" & TRIM$(Str$(NROCOT&)) & ";" & TRIM$(CRANGOEXT$)
    Call STDFORM("OCMA-CTP", FILTX$)
  'POLLERA
    CODICOTIZADO = CodigodeCotizacion&(NROCOT&, "POLLERA")
    CRANGOEXT$ = CODEXT$(CODICOTIZADO)
    If TRIM$(CRANGOEXT$) = "" Then CRANGOEXT$ = "-----"
    FILTX$ = TRIM$("102030-01") & ";" & TRIM$(Str$(NROFTEC&)) & ";" & TRIM$(Str$(NROCOT&)) & ";" & TRIM$(CRANGOEXT$)
    Call STDFORM("OCPO-CTP", FILTX$)



End Sub

Private Sub Command49_Click()
            Call CARGAMSF1(1)
            Call CARGAMSF2(1)
            Call CARGAMSF3(1)
            Call CARGAMSF4(1)
            Call CARGAMSF5(1)


End Sub

Private Sub Command5_Click()
   X = LISTASELECGENERAL("TABLASI_NO", , , 1)
   Me.TXTIMPRESION = TRIM$(RESP_ZELE_VECT(2))
   Call VALORESPORDEFECTO

End Sub


Private Sub Command50_Click()
        Command50.Enabled = False
        FTEC& = XVALO(Text5)
        If FTEC& < 1 Then
          Call COMUNI("Imposible Continuar No Hay una Ficha Técnica Activa")
          GoTo 99
        End If
        
'        'debe estar grabada la ficha tecnica.
'        If CantRegistros&("RECETACTP", "NUMRECETACTP = " & FTEC&) < 1 Then
'         Call COMUNI("Debe Primero Grabar la Ficha Técnica")
'         GoTo 99
'        End If
        '
10 '     OPX% = COMALTER%("Agregar + Datos\\Cancelar\Datos Anexos\Hilachas cauterizadas y-o Uso Detector Metales") LAUTARO PIDIO QUE SE QUITE LA OPCION
'        If OPX% < 2 Then GoTo 99
'        If OPX% = 3 Then
         CotiMasDato.Show 1
         Exit Sub
'           GoTo 10
'        End If
        
        'SI NO ES NINGUNA DE LAS OPCIONES ANTERIORES SIGUE CON EL CODIGO
        If DERACCE("MODDTANE", "Modificar Datos Anexos") < 2 Then GoTo 99
        Unload VENTABNEW
        VENTANA.Inicializar = 0
        VENTANA.CampEditables = "1;3"
        VENTANA.AgregaRegistro = 1
        VENTANA.UtilizaFormula = 1
        VENTANA.HabilitaBorrar = 1
        VENTANA.HabilitaInsertar = 1
        VENTANA.NoMostrarBotonTilde = 0
        VENTANA.TITULO = "Agregar Datos Anexos"
        VENTANA.ColEcoCampo(1) = "CAMPO1;CAMPO2;DATASQL;LISTA = 'DATOSLIB'" 'ESTA ES LA TABLA
        VENTANA.Inicializar = 1
        Campos$ = "Rótulo/A48;" '1
        Campos$ = Campos$ + "Texto/A48;"       '2
        Campos$ = Campos$ + "Valor/A12;"       '3
        Campos$ = Campos$ + "Ftecnica/A0"     '4
        VENTANA.Campox = Campos$
        
        Dim obj As New RECORXET
        CAMPOSTABLA$ = "CAMPO1 ,CAMPO2 ,CAMPO3 ,CAMPO4"
        TABLA$ = "DATASQL "
        CONDI$ = " CAMPO4  = '" & CStr(FTEC&) & "' AND LISTA = 'DATOSEXT'" 'ESTA ES LA LISTA CON EL IDENTIFICADOR QUE SE VAN A GUARDAR
        '
        Set RS1 = obj.RS_GENERAL(TABLA$, CAMPOSTABLA$, CONDI$)
        Call Carga_MSF_Recordset(RS1, Campos$, VENTABNEW.MSF)
        On Error Resume Next
        Set obj = Nothing
        RS1.Close
        Set RS1 = Nothing
        On Error GoTo 0
        '
20      VENTABNEW.Show 1
        If VENTABNEW.APROBADO% < 1 Then GoTo 99


        'TRANSACCIÓN GRABACION
        On Error GoTo ERROR_GUARDAR
        FLEXCONN.BeginTrans
        CONDI$ = " CAMPO4  = '" & CStr(FTEC&) & "' AND LISTA = 'DATOSEXT'"
        Call ACTUREGISTRO("DATASQL", " MARBORRA ", CONDI$, 1, REGAF&, "NOMESS")
        SQLX$ = "SELECT * FROM DATASQL"
        SQLX$ = SQLX$ + " WHERE CAMPO1 = ''"
        
        Dim MSF As msFlexGrid
        Set MSF = VENTABNEW.MSF
        
        Dim RS As ADODB.Recordset
        Set RS = New ADODB.Recordset
        RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText

        With RS
              For REG& = 1 To MSF.Rows - 1
                  If TRIM$(MSF.TextMatrix(REG&, 1)) = "" Then GoTo 50
                  .AddNew
                  !CAMPO1 = MSF.TextMatrix(REG&, 1)
                  !campo2 = MSF.TextMatrix(REG&, 2)
                  !CAMPO3 = MSF.TextMatrix(REG&, 3)
                  !CAMPO4 = CStr(FTEC&) 'ESTOS DOS CAMPOS SON LOS IDENTIFICADORES
                  !Lista = "DATOSEXT"   'ESTOS DOS CAMPOS SON LOS IDENTIFICADORES
                  !MARBORRA = 0
                  .Update
50            Next REG&
        End With
        
        RS.Close
        Set RS = Nothing
        
        CONDI$ = CONDI$ + " AND  MARBORRA = 1  "
        Call BORRAREGISTROS("DATASQL", CONDI$, REGAF&, "NOMESS")
        Set VENTANA = Nothing
        
ERROR_GUARDAR:
        If Err <> 0 Then
90         FLEXCONN.RollbackTrans
           On Error GoTo 0
           Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION")
        Else
          FLEXCONN.CommitTrans
          Call COMUNI("Grabación Realizado Exitosamente")
        End If
        Unload VENTABNEW
       
99     Command50.Enabled = True
End Sub

Private Sub Command51_Click()
              ANCHO = XVALO(TXTANCHOCM)
              Largo = XVALO(TXTCUERPO_LARGOCM)
              ALTO = XVALO(TXTCUERPOALTOCM)
'**** se quita validacion (incluida por raul) a soliictud de Franco ****
'****'****'****'****'****'****'****'****'****'****'****'****'****'****'****
'              If ANCHO = LARGO Then 'SOLO EN LOS CASOS QUE ANCHO SEA IGUAL A LARGO
                    'CARGAR DATOS DE CORTE - LEO LAS FORMULAS
                    SQLX$ = "SELECT TAPAPOLLERA,POLLERA,CUERPOLAT,FONDOPANEL,LATERAL2,UPANEL "
                    SQLX$ = SQLX$ + " FROM CALCUCORTE WITH(NOLOCK)  WHERE CODBIGBAG = '" & TRIM$(TXTCODCONFECCION) & "'"
                    Set RS = READSET(SQLX$)
                    With RS
                      If Not .EOF Then
                          formulaTapa$ = SAFETEXT$(!TAPAPOLLERA)
                          formulaPollera$ = SAFETEXT$(!Pollera)
                          formulaCuerpo$ = SAFETEXT$(!CUERPOLAT)
                          formulaLateral$ = SAFETEXT$(!LATERAL2)
                          formulaFondo$ = SAFETEXT$(!FONDOPANEL)
                          formulaUPanel$ = SAFETEXT$(!UPANEL)
                      End If
                    End With
                    RS.Close: Set RS = Nothing
                    '
                    'ASIGNO LOS RESULTADOS DE LAS FORMULAS A CADA FILA EN EL CAMPO CORTE DE LA GRILLA
                    MSF1.TextMatrix(1, 6) = CORTE_CONTEPLAST(formulaTapa$, ALTO, ANCHO, Largo, "TAPA", "F12.2")
                    MSF1.TextMatrix(2, 6) = CORTE_CONTEPLAST(formulaPollera$, ALTO, ANCHO, Largo, "POLLERA", "F12.2")
                    MSF1.TextMatrix(3, 6) = CORTE_CONTEPLAST(formulaCuerpo$, ALTO, ANCHO, Largo, "CUERPO", "F12.2")
                    MSF1.TextMatrix(4, 6) = CORTE_CONTEPLAST(formulaLateral$, ALTO, ANCHO, Largo, "LATERALES", "F12.2")
                    MSF1.TextMatrix(5, 6) = CORTE_CONTEPLAST(formulaFondo$, ALTO, ANCHO, Largo, "FONDO", "F12.2")
                    MSF1.TextMatrix(6, 6) = CORTE_CONTEPLAST(formulaUPanel$, ALTO, ANCHO, Largo, "UPANEL", "F12.2")
                    '
                    'COPIO EN LA ULTIMA COLUMNA NO VISIBLE PARA VALIDAR LUEGO SI SE CAMBIO ALGUNO
                    MSF1.TextMatrix(1, 10) = MSF1.TextMatrix(1, 6)
                    MSF1.TextMatrix(2, 10) = MSF1.TextMatrix(2, 6)
                    MSF1.TextMatrix(3, 10) = MSF1.TextMatrix(3, 6)
                    MSF1.TextMatrix(4, 10) = MSF1.TextMatrix(4, 6)
                    MSF1.TextMatrix(5, 10) = MSF1.TextMatrix(5, 6)
                    MSF1.TextMatrix(6, 10) = MSF1.TextMatrix(6, 6)
                    For i& = 1 To MSF1.Rows - 1 'recalculo coMO en el ventabnew
                        MSF1.TextMatrix(i&, 7) = FORMATNUM$(XVALO(MSF1.TextMatrix(i&, 5)) * XVALO(MSF1.TextMatrix(i&, 6)), "F12.4")
                        MSF1.TextMatrix(i&, 9) = FORMATNUM$((XVALO(MSF1.TextMatrix(i&, 7)) * XVALO(MSF1.TextMatrix(i&, 8))) / 100, "F12.4")
                    Next i&

'              End If

End Sub

Private Sub Command52_Click()
    FRMCOLOR.LICHESEL.Clear
    
    SQLX$ = "SELECT * FROM DATASQL WITH(NOLOCK) WHERE LISTA = 'COLORES'"
    Set RS = READSET(SQLX$)
    With RS
      Do While Not .EOF
        COD1% = XVALO(!CAMPO1)
        COLORX$ = SAFETEXT$(!campo2)
        Call REPLA(TXTX$, SAFETEXT$(COD1%), 1, 5)
        Call REPLA(TXTX$, SAFETEXT$(COLORX$), 7, 24)
        FRMCOLOR.LICHESEL.AddItem TXTX$
        .MoveNext
      Loop
    End With
    FRMCOLOR.Show 1
    If TRIM$(FRMCOLOR.COLOR) <> "" Then
       Text32 = TRIM$(FRMCOLOR.COLOR)
    End If
End Sub

Private Sub Command6_Click()
   X = LISTASELECGENERAL("CTP.CARAS", , , 1)
   Me.TXTCANTICARAS = TRIM$(RESP_ZELE_VECT(2))
   Call VALORESPORDEFECTO

End Sub

Private Sub Command7_Click()
   X = LISTASELECGENERAL("CTP.CANTIUSOS", , , 1)
   Me.TXTCANTIUSOS = TRIM$(RESP_ZELE_VECT(2))

End Sub

Private Sub Command8_Click()
   X = LISTASELECGENERAL("CTP.PAB", , , 1)
   Me.TXTTELAVENTILADA = TRIM$(RESP_ZELE_VECT(2))
End Sub

Private Sub Command9_Click()
   X = LISTASELECGENERAL("CTP.GRAMAJE", , , 1)
   Me.TXTGRAMAJECUERPO = TRIM$(RESP_ZELE_VECT(2))
   Call VALORESPORDEFECTO

End Sub


Private Sub Form_Load()
    SSTab1.Tab = 0
    '22092020
    If App.EXEName = "COTIZACO" Or App.EXEName = "COTIZA99" Then
      If (App.PrevInstance = True) Then
          MsgBox "Yá se Encuentra Abierto el Módulo de Ficha Técnica"
          Call FINAL("")
      End If
      
      Call VERJOBID(OKEY%)
      If OKEY% < 1 Then Call FINAL("")
      Call ABRECONEXION
      Call CONFIGMENU
      Label8.Visible = True
      
      Dim CTP As New Conteplast
      Call CTP.ACTUBASEWALDBOTT
      Set CTP = Nothing
      
      Label8.Visible = False
      Call CREATABLARECETACTP
      Call MIGRADAT_A_SQL
    Else
'      Me.Visible = False
    End If
    Call GENERAR_ESTRUCTURAS_SQL
    Call CREARTABLASOLICOTI
    ActualizaListaSeleccion = True
    LBLNUEVAFICHATECNICA.Visible = True
    TXTFECHA = FECHATEX$(HOY(HO$))
    Text5 = SAFETEXT$(NumNuevaReceta)
    Picture5.Top = 0
    VScroll1.Min = 0
    VScroll1.Max = (Picture5.Height - Picture6.Height) / 100
    Call BLANFORMU
    Call CARGAMSF1(0)
    Call CARGAMSF2(0)
    Call CARGAMSF3(0)
    Call CARGAMSF4(0)
    Call CARGAMSF5(0)
    Call FRMCOTI2.CARGATEXTBLOQUEETIQUETAS(0)
    Call FRMCOTI2.CARGAROTULOSGASTOS(1)
    '
    'INGRESO DE VENDEDOR SEGUN USUARIO
    USNU% = USNBR% Mod 100
    
    If USNU% > 1 Then
       VENDESEGUNUSUARIO% = LOGVEND.VendeVsUsuario(USNU%)
       If VENDESEGUNUSUARIO% < 1 Then ' SI NO SON VENDEDORES DEBEN ASIGNARSE VENDEDOR 1 GERENCIA.
            LOGVEND.Text1 = USNU%
            LOGVEND.Label1 = USERNAME$
            LOGVEND.Show 1
       Else
          Text4 = VENDESEGUNUSUARIO%
       End If
    End If
    MODIFIC% = 0
End Sub

Sub CARGAMSF5(CARGARDATOS%)
           Campos$ = "COMPONENTE/A20;FICHA TECNICA(g-m2)/A20;CODIGO/A16;DESCRIPCION/A24;MODULO/F12;CORTE/F10.2;CONSUMO (cm)/f12.2;PRECIO (u$s*m)/F12.2;PRECIO (u$s)/F12.2"
           Call CARGA_ENCABEZADO(MSF5, Campos$, Me)
           MSF5.Rows = 2
           MSF5.TextMatrix(1, 1) = "ESLINGA PERIMETRAL [CM]"
           
           If CARGARDATOS% > 0 Then
              MSF5.TextMatrix(1, 2) = TXTESLINGAPERIMETRAL
              '
              'COMPLETO LA COLUMNA MODULO DE LA TABLA CONFIGBB
              'CINTDESC, CINTDESCSEG,CINTCARGPOLL, CINTCARGSEG,  CINTLINER,
              SQLX$ = "SELECT MAMPARO,ESLINGA,ESLPERIM "
              SQLX$ = SQLX$ + " FROM CONFIGBB WITH(NOLOCK)  WHERE CODBIGBAG = '" & TRIM$(TXTCODCONFECCION) & "'"
              Set RS = READSET(SQLX$)
              With RS
                If Not .EOF Then
                    MSF5.TextMatrix(1, 5) = SAFETEXT$(!ESLPERIM)
                End If
              End With
              RS.Close: Set RS = Nothing
                 RESUL# = (XVALO(TXTANCHOCM) + XVALO(TXTCUERPO_LARGOCM)) * 2 + 25
              MSF5.TextMatrix(1, 6) = TRIM$(FORMATNUM$(RESUL#, "F12.2"))
           End If
End Sub


Sub MIGRADAT_A_SQL()

    Call CREA_TABLA_USUARIOSSQL
    Call ACTUALIZARFLEXUSR
    '''''''''''''''''''''''EL IDENTIFICADOR Y LA CLAVE PARA ASIGNARLA MAS ABAJO
    FUSER$ = LOADFILE$("USER.DAT")
    Dim IDENTIFICADOR1$()
    Dim CLAVEUSERDAT1$()
    For II% = 1 To Len(FUSER$) Step 64
       jj% = jj% + 1
       F0$ = Mid$(FUSER$, II%, 64)
       CLAV$ = Mid$(F0$, 43, 12): If Left$(CLAV$, 1) = "@" Then CLAV$ = DESENCRI$(Mid$(F0$, 44, 11)) + " "
       ReDim Preserve CLAVEUSERDAT1$(jj%), IDENTIFICADOR1$(jj%)
       CLAVEUSERDAT1$(jj%) = TRIM$(CLAV$)
       Ident$ = Mid$(F0$, 31, 12)
       If Left$(Ident$, 1) = "@" Then Ident$ = DESENCRI$(Mid$(Ident$, 2, 11)) + " "
       IDENTIFICADOR1$(jj%) = TRIM$(Ident$)
    Next II%
    '
    'tomo el nivel de configuracion y lo cargo en un vector para ganar velocidad ya que la funcion demora mucho
    Dim NICON1%()
    For i& = 1 To ULTREG&("DATUSERS")
       ReDim Preserve NICON1%(i&)
       X$ = REGLEIDO$("DATUSERS", i&)
       NICON1%(i&) = Asc(Mid$(X$, 32, 1))
    Next i&
    '
    'LEO EL IDMAQUINA QUE SE GUARDA EN LA TABLA SQL
    SQLX$ = "SELECT IDMAQUINA FROM USUARIOSSQL WITH(NOLOCK)"
    Set RS1 = READSET(SQLX$)
    Dim CODIMAQUS$()
    L% = 0
    With RS1
      Do While Not .EOF
        L% = L% + 1
        ReDim Preserve CODIMAQUS$(L%)
        CODIMAQUS$(L%) = SAFETEXT$(!IDMAQUINA)
        .MoveNext
      Loop
    End With
    ''''''''''''''''''''''''''''''''' HASTA AQUI
    '
    Call PoneEnCeroNull("USUARIOSSQL", "MARBORRA", REGAF&, "NOMESS")
    Call ACTUREGISTRO("USUARIOSSQL", "MARBORRA", "MARBORRA < 1", 1, REGAF&, "NOMESS")
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
25  On Error Resume Next
    SQLX$ = "SELECT * FROM USUARIOSSQL WHERE NUSER < 1 "
    RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    '
    FE% = HOYSERVER%
    For i& = 1 To ULTREG&("FLEXUSR")
       X$ = REGLEIDO$("FLEXUSR", i&)
       NUS% = Asc(Mid$(X$, 1, 1))
       DENOM$ = Mid$(X$, 2, 29)
       NICON% = NICON1%(NUS%)
       RS.AddNew
       RS!FECHA = CVDAT(FE%)
       RS!NUSER = NUS%
       RS!CLAVE = CLAVEUSERDAT1$(i&) 'CLAVEUSERDAT$(NUS%, RET_IDENTIFICADOR$)
       RS!nombre = TRIM$(DENOM$)
       RS!IDENTIFICADOR = IDENTIFICADOR1$(i&) 'RET_IDENTIFICADOR$ ' VALOR DE RETORNO
       RS!NIVELCONF = NICON%
       RS!IDMAQUINA = CODIMAQUS$(NUS%) ' VECTOR CARGADO MAS ARRIBA
       RS!MARBORRA = 0
       RS.Update
    Next i&
    RS.Close
    Set RS = Nothing
    Call BORRAREGISTROS("USUARIOSSQL", "MARBORRA = 1", REGAF&, "NOMESS")
End Sub
Private Sub mnuAltaNuevaVersion_Click()
    If TRIM(TXTCODIGO) = "" Then
       Call COMUNI("Atención !!! Debe Haber un Código Activo")
       Exit Sub
    End If
    CII% = CODINT%(TXTCODIGO)
    NUEVOCODI.TXTCLIENTE = XVALO(NCLIE)
    Erase CADENATEX$
    If NUEVAVERSION% > 0 Then
       NUEVAVER$ = GENERAR_NUEVA_VERSION$(CII%)
       If NUEVAVER$ = "" Then NUEVAVERSION% = 0 'NO PUDO GENERAR LA NUEVA VERSION
    End If
    Call TEXTOLOTES(NUEVAVER$, "-")
    On Error Resume Next
    For i& = 1 To UBound(CADENATEX$)
         Select Case i&
           Case 1: NUEVOCODI.TXTTIPOBOLSA = CADENATEX$(i&)
           Case 2: NUEVOCODI.TXTCLIENTE = CADENATEX$(i&)
           Case 3: NUEVOCODI.TXTORDEN = CADENATEX$(i&)
           Case 4: NUEVOCODI.TXTVERSION = CADENATEX$(i&)
         End Select
    Next i&
    On Error GoTo 0
20  NUEVOCODI.Show 1
    ESPORCOPIA% = 0
    If NUEVOCODI.TXTCODIGO = "" Then Exit Sub
    
       VDEF$ = String$(128, 0)
       '
       'PASA LOS DATOS DEL
       LONCODI% = XVALO(Control$("MASTER", "LOMAXCOD"))
       COD$ = TRIM$(NUEVOCODI.TXTCODIGO)
       EXISTECOD% = CODINT%(COD$)
       If EXISTECOD% > 0 Then
          Call COMUNI("Imposible dar de Alta - Código Existente")
          GoTo 20
       End If
       CI_NEW% = CODINT_NEW
       Call REPLA(VDEF$, MKI$(CI_NEW%), 1, 2)
       If CII% > 0 Then 'ES POR COPIA
        ESPORCOPIA% = 1
        Call REPLA(VDEF$, MKI$(GRUCOVEI%(CII%)), 3, 2)
        Call REPLA(VDEF$, DESCRIT0$(CII%), 5, 64)
        Call REPLA(VDEF$, COD$, 77, LONCODI%)
        Call REPLA(VDEF$, Unimed$(CII%), 109, 2)
        Call REPLA(VDEF$, Chr$(5), 101, 1)
       Else
        Call REPLA(VDEF$, COD$, 77, LONCODI%)
        Call REPLA(VDEF$, "U.", 109, 2)
        Call REPLA(VDEF$, Chr$(1), 101, 1)
       End If
       '
       OKX% = 0
       
       Call ALTACODIGO(VDEF$, OKX%)
       
       If LONCODI% < 15 Or LONCODI% > 24 Then LONCODI% = 15
       If OKX% = 1 Then
          CODINUE$ = Mid$(VDEF$, 77, LONCODI%)
       End If
       '
       Screen.MousePointer = 11
       Call BLOQUEAR_CONTROLES(Me, 2, 1, 1, 1)
       NOCERRARFRM% = 1 ' NO PERMITE CERRAR EL FORMULARIO
       If OKX% = 1 And ESPORCOPIA% > 0 Then
        Dim RMaster As ADODB.Recordset
        Set RMaster = New ADODB.Recordset
        strSQL$ = "Select * from Master WITH(NOLOCK) where codint=" & CII%
        Set RMaster = READSET(strSQL$)
        If Not (RMaster.EOF And RMaster.BOF) Then
           CACAM% = RMaster.Fields.Count
           For KU% = 1 To CACAM%
             Dato$ = SAFETEXT$(RMaster.Fields(KU% - 1).Value)
             NOCAMPO$ = TRIM$(UCase$(RMaster.Fields(KU% - 1).Name))
             If NOCAMPO$ <> "CODIGO" Then
                If NOCAMPO$ <> "DESCRIPCION" Then
                   If NOCAMPO$ <> "CODINT" Then
                      If NOCAMPO$ <> "ALIAS" Then
                         If NOCAMPO$ <> "UMEDIDA" Then
                           CONDI$ = "CODINT=" & CI_NEW%
                           Call ACTUREGISTRO("MASTER", NOCAMPO$, CONDI$, Dato$, REGAF&, "NOMESS")
                         End If
                      End If
                   End If
                End If
              End If
           Next KU%
        End If
       End If
       'ACTUALIZO EL MASTER.
       CONDI$ = " CODIGO <> '' AND STATUS > 0 ORDER BY CODIGO ASC"
       If Control$("", "IDSECUN") = "ALIAS" Then
           Call CARGALISEL("ECOMAST", "MASTER", "CODIGO/A24;DESCRIPCION/A36;NPLANO/A24", CONDI$, "NOMESS")
       Else
           Call CARGALISEL("ECOMAST", "MASTER", "CODIGO/A24;DESCRIPCION/A50", CONDI$, "NOMESS")
       End If
        '
        'Call FRMENVAP.COPIAR_ESTRUCTURA(CII%, CI_NEW%)

        On Error Resume Next
        Unload NUEVOCODI
        On Error GoTo 0
        
'        NCLIEN = XVALO(Mid$(COD$, 4, 5))
'        FRMENVAP.TXTOBS(31) = NCLIEN
'        FRMENVAP.TXTCODIGO = COD$
'        FRMENVAP.TXTRAZOCLI = rasocli$(NCLIEN)
'        FRMENVAP.TXTPREPARO = USERNAME$
'        FRMENVAP.TXTFECHPREPARO = FECHATEX$(HOY(HO$))
'        Screen.MousePointer = 1
        Call BLOQUEAR_CONTROLES(Me, 1, 1, 1, 1)
'        NOCERRARFRM% = 0

End Sub

Private Sub missolicitudesAprobacionft_Click()
  NUS% = USNBR% Mod 100
  VENDE% = LOGVEND.VendeVsUsuario(NUS%)
  NRO = SELEC_RECETA(0, 2, , -1, VENDE%)
  If NRO > 0 Then
     Text5 = NRO
  End If
  
   'SIEMPRE QUEDA HABILITADA LA POSIBILIDAD DE AGREGAR UN ARCHIVO
   Command47.Enabled = True
   Command43.Enabled = True
   Command45.Enabled = True
   Combo2.Enabled = True

End Sub

Private Sub mnuAbrirCotExistente_Click()
    Dim CTP As New Conteplast
    LBLNROSOLICITUDCOTIZACION = ""
    CONDI$ = " AND (FT_STAT < 9 OR FT_STAT IS NULL) AND COT_NROCOT > 0 AND  (COT_NoMostEnList < 1 OR COT_NoMostEnList  IS NULL) "
    '
    NROFT = CTP.SELEC_FTECNICAYCOTI(CONDI$)
    NROCOT& = XVALO(RESP_ZELE_VECT(3))
    
    If FRMCOTI2.STATUSCOTIZA(NROCOT&) = 1 Then
        Call COMUNI("Cotización Aprobada La Misma, Deshabilitada Para Registrar Cambios")
        FRMCOTI2.Command49.Enabled = False
        FRMCOTI2.Command47.Enabled = False
        Command46.Enabled = False
'        mnuGuardarCoti.Enabled = False
    Else
        FRMCOTI2.Command49.Enabled = True
        FRMCOTI2.Command47.Enabled = True
        Command46.Enabled = True
        mnuGuardarCoti.Enabled = True
    End If
    If NROFT > 0 Then
     lblNuevaCotizacion.Visible = False
     Text5 = ""
     Text5 = NROFT
     TXTNUMCOTIZACION = NROCOT&
     LBLNROSOLICITUDCOTIZACION = NROSOLICITUDCOTIZACION(NROCOT&)

     Call FRMCOTI2.CARGAROTULOSGASTOS(1)
     Call VALORESPORDEFECTO
     Call FRMCOTI2.VALORESCALUCLADOSBLOQUEETIQUETAS
     Call FRMCOTI2.IMPORTESSBLOQUEETIQUETAS
     Call FRMCOTI2.CALCULARCOSTOMATERIALES
     MODIFICCOT1% = 0
     MODIFICCOT2% = 0
     SSTab1.Tab = 2
     Command46.Enabled = True
'     Call CARGAROTULOSGASTOS
    End If

End Sub

Private Sub mnuActualizarBaseClientes_Click()
   ARCHIVOWALDBOT$ = Control("COTICONT", "ARCDROPB") & "\SGIW14.MDB"
   If EXISTE%(ARCHIVOWALDBOT$) < 1 Then
      OPX% = COMALTER%("Base de Clientes No Encontrada (" & ARCHIVOWALDBOT$ & ")\\Abandonar\Forzar Activación de DropBox y Volver a Intentar")
      If OPX% = 2 Then
           XYZ = Shell("X:\FLEXOFT\ArchivoBatActivarDropBox\ABRIRDROPBOX.bat", vbMinimizedNoFocus)
           Screen.MousePointer = 11
           Call T_Espera(40)
           Screen.MousePointer = 1
           GoTo 20
      End If
      Exit Sub
   End If
   
20 Label8.Visible = True
   Dim CTP As New Conteplast
   Call CTP.ACTUBASEWALDBOTT
   Set CTP = Nothing
   Label8.Visible = False
   
End Sub

Private Sub mnuAnularCotización_Click()

    'MUESTRA INFORMACINO DE BD EN BASE A UNA VISTA DE LAS TABLAS RECETACTP (F.TECNICA) Y LA TABLA DE COTIZACIONES
    Dim CTP As New Conteplast
    CONDI$ = "AND (FT_STAT < 9 OR FT_STAT IS NULL)"
    NRO = CTP.SELEC_FTECNICAYCOTI(CONDI$)
    NROCOT& = XVALO(RESP_ZELE_VECT(4))
    If XVALO(RESP_ZELE_VECT(4)) > 0 Then
       Call COMUNI("Imposible Anular\La Ficha Técnica Seleccionada Contiene Cotización N°: " & NROCOT&)
       Exit Sub
    End If
    
    If NRO > 0 Then
     CONDI$ = " NUMRECETACTP = " & NRO
     Call ACTUREGISTRO("RECETACTP", "STATUS", CONDI$, 9, REGAF&, "NOMESS")
     If XVALO(Text5) = NRO Then
       Text5 = ""
       Call mnuNuevaReceta_Click
       Call VALORESPORDEFECTO
     End If
    End If

    
    
End Sub

Private Sub mnuBajaCotización_Click()
    If DERACCE%("BajaSoCt", "Baja de Soicitud de Cotización") < 2 Then Exit Sub
    
    NRO = NROSOLICITUD(1, 0)
    'BORRA SOLICITUD DE COTIZACION
    If NRO < 1 Then Exit Sub
    On Error GoTo ERROR_GUARDAR
    FLEXCONN.BeginTrans
    SQLX$ = " DELETE FROM SOLICOTI WHERE SOLICOTI_ID  = " & NRO
    FLEXCONN.Execute (SQLX$)
    
ERROR_GUARDAR:
    If Err <> 0 Then
       MsgBox Err.Description
       On Error Resume Next
       FLEXCONN.RollbackTrans
       On Error GoTo 0
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION")
       Screen.MousePointer = 1
       Exit Sub
    Else
      FLEXCONN.CommitTrans
      Call COMUNI("Grabación Realizado Exitosamente")
     End If
     
End Sub

Private Sub mnuBajaFichaTecnica_Click()
   
   If FIRMAELECTRONICAOK% <= 0 Then Exit Sub
   NROFTECNICA& = SELEC_RECETA()
   If NROFTECNICA& < 1 Then Exit Sub
   CODIGOMASTER$ = TRIM$(RESP_ZELE_VECT(5))
   
   'VALIDAR SI NO TIENE CODIGO SE BORRA
   CI1% = CODINT%(CODIGOMASTER$)
   If CODIGOMASTER$ = "" Then 'directamente elimina la Ficha Técnica
      'directamente elimina la Ficha Técnica
      GoTo 30
      
   ElseIf CODINT%(CODIGOMASTER$) < 1 Then
      Call COMUNI("Código No Válido o Dado de Baja\Deberá Darlo de Alta y Volver a Intentar")
      Exit Sub
   End If
   '
   TIENEPEDIDO% = CantRegistros("PEDDETA", "CODIINT = " & CI1%, "NOMESS")
   If TIENEPEDIDO% > 0 Then
      Call COMUNI("Imposible Eliminar Ficha Técnica el Mismpo Posee Pedido")
      Exit Sub
   End If
   
   TIENEPRESUPUESTO% = CantRegistros&("PRESUVEN", "CODIINT = " & CI1%, "NOMESS")
'   NROPRESUPUESTO& = XVALO(VALOBADA("", "NroPres", , "NOMESS"))
   'PUEDE TENER COTIZACION APROBADA PERO NO HABERSE GENERADO PEDIDO
   TIENECOTIZACIONAPROBADA% = XVALO(VALOBADA("COTIZACTP", "STATUS", "NUMRECETACTP = " & NROFTECNICA&, "NOMESS"))
   NROCOTIZACION& = XVALO(VALOBADA("COTIZACTP", "NROCOTIZACION", "NUMRECETACTP = " & NROFTECNICA&, "NOMESS"))
   
   TXTX1$ = "Confirmación de Eliminación Permanente de Ficha Técnica: " & NROFTECNICA&
   TXTX1$ = TXTX1$ + "\Tenga en Cuenta que Si la Ficha Técnica Posee Cotización (Aprobada o No) "
   TXTX1$ = TXTX1$ + "\Se Eliminarán Conjuntamente con La Ficha Técnica"
   TXTX1$ = TXTX1$ + "\Si Posee Código Generado el Mismo Se Dará de Baja"
   TXTX1$ = TXTX1$ + "\\Cancelar\Eliminar"
    
    OPX% = COMALTER%(TXTX1$)
    If OPX% < 2 Then Exit Sub
   
       'TRANSACCIÓN GRABACION
30     On Error GoTo ERROR_GUARDAR
       FLEXCONN.BeginTrans

      'BORRA SOLICITUD DE COTIZACION
       SQLX$ = " DELETE FROM SOLICOTI WHERE CODIMASTER = '" & CODIGOMASTER$ & "'"
       FLEXCONN.Execute (SQLX$)
       'BORRA ENCABEZADO DE COTIZACION
       SQLX$ = " DELETE FROM COTIZACTP WHERE CODIMASTER = '" & CODIGOMASTER$ & "' AND NUMRECETACTP = " & NROFTECNICA&
       FLEXCONN.Execute (SQLX$)
       'BORRA DETALLE DE COTIZACION 1
       SQLX$ = " DELETE FROM COTI1 WHERE CODINTCONJ = " & CI1% & " AND NROCOTIZACION = " & NROCOTIZACION&
       FLEXCONN.Execute (SQLX$)
       'BORRA COTIZACION 2 (VALORES)
       SQLX$ = " DELETE FROM COTI2 WHERE NROCOTIZACION = " & NROCOTIZACION&
       FLEXCONN.Execute (SQLX$)
       'BORRA FICHA TECNICA
       SQLX$ = " DELETE FROM RECETACTP WHERE NUMRECETACTP = " & NROFTECNICA&
       FLEXCONN.Execute (SQLX$)
       'BORRA PRESUPUESTO EN PRESUVEN
       SQLX$ = " DELETE FROM PRESUVEN WHERE  CODIINT = " & CI1%
       FLEXCONN.Execute (SQLX$)
       'BAJA DE CODIGO
       'SI EL CODIGO ES PARTE DE UNA ESTRUCTURA NO SE DARA DE BAJA , PERO SI EL RESTO ANTERIOR
       Call BAJAARTICULO(CI1%)
       
      '
   
   
ERROR_GUARDAR:
    If Err <> 0 Then
       A = Err.Description
       MsgBox A
       On Error Resume Next
       FLEXCONN.RollbackTrans
       On Error GoTo 0
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION")
       Screen.MousePointer = 1
       Exit Sub
    Else
      FLEXCONN.CommitTrans
      Call COMUNI("Grabación Realizado Exitosamente")
     End If
     
   
'   TIENECOTIZACION% = CantRegistros("COTI1", "CODINTCONJ = " & CI1%, "NOMESS")
'
''   TIENEPEDIDO% = CantRegistros("PEDDETA", "CODIINT = " & CI1%, "NOMESS")
'
'
'   TXTNROPED = NroPedidoSegunOf&(NORFA$)
'   TXTNROCOTIZA = NROCOTIZA(XVALO(TXTNROPED))
'   NROFT& = CTP.NroFTecnicaXCodigo(CI1%)

   'SI TIENE CODIGO BORRA LA RECETA Y SE DA DE BAJA EL CODIGO
   'SI NO TIENE PEDIDO Y ORDEN DE FABRICACION O NO FIGURA EN MOVISTO
   'SI TIENE COTIZACION Y ESTA EN PRESUVEN Y NO TIENE PEDIDO HABRIA QUE BORRAR TAMBIEN DE PRESUVEN
   'SI  TIENE PEDIDO U  ORDEN DE FABRICACION NO SE PUEDE ANULAR
   'SI TIENE COTIZACION SI TIENE COTIZACION SE DEBERIA DAR DE BAJA A LA COTIZACION SI NO TIENE PEDIDO
   'VALIDAR SI TIENE PEDIDO
   'VALIDAR SI SE TIENE CODIGO Y ORDEN DE FABRICACION
   

End Sub

Function FIRMAELECTRONICAOK%()
     FIRMAELEC.Text1 = ""
     FIRMAELEC.Text2 = ""
     FIRMAELEC.Label1 = ""
     FIRMAELECTRONICAOK% = -1
     FIRMAELEC.Show 1

'     If FIRMAELEC.ESTADO > 0 Then
'        FIRMAELECTRONICAOK% = 1
'        Exit Function
     If FIRMAELEC.ESTADO < 1 Then
        FIRMAELECTRONICAOK% = -1
        Exit Function
     Else
        If NIVCONFI%(XVALO(FIRMAELEC.Label1)) < 3 Then
         Call COMUNI("Usuario Con Nivel de Clasificación Inválido Para Esta Operación")
         FIRMAELECTRONICAOK% = -1
         Exit Function
        End If
     End If

End Function
Private Sub mnuBB_Click()
        If DERACCE("CONBIGBA", "CONFIGURACION DE BIG BAG") < 2 Then Exit Sub
        Unload VENTABNEW
        VENTANA.Inicializar = 0
        VENTANA.CampEditables = "1;3;4;5;6;7;8;9;10;11;12;13;14;15;16;17;18;19;20;21;22;23;24"
        VENTANA.AgregaRegistro = 1
        VENTANA.UtilizaFormula = 1
        VENTANA.HabilitaBorrar = 1
        VENTANA.HabilitaInsertar = 1
        VENTANA.NoMostrarBotonTilde = 0
        VENTANA.TITULO = "CONFIGURACIÓN DE  BIGBAG"
        VENTANA.ColEcoCampo(1) = "CAMPO1;CAMPO2;DATASQL;LISTA  = 'CODBB'"
        VENTANA.Inicializar = 1
        Campos$ = "CODIGO/A10;" '0
        Campos$ = Campos$ + "DESCRIPCION/A48;"      '1
        Campos$ = Campos$ + "CUERPO C.C./A11;"      '2
        Campos$ = Campos$ + "LATERALES C.C./A11;"   '3
        Campos$ = Campos$ + "FONDO C.C./A11;"       '4
        Campos$ = Campos$ + "TAPA C.C./A11;"        '5
        Campos$ = Campos$ + "POLLERA C.C./A11;"     '6
        Campos$ = Campos$ + "U PANEL C.C./A11;"     '7
        Campos$ = Campos$ + "VALV.CARGA/A11;"       '8
        Campos$ = Campos$ + "VALV.DESC/A11;"        '9
        Campos$ = Campos$ + "CINTILLA POLLERA/A16;"    '10
        Campos$ = Campos$ + "CINTILLA DESC/A13;"    '11
        Campos$ = Campos$ + "CINTILLA DESC. SEG/A13;" '12
        Campos$ = Campos$ + "CINTILLA CARGA POLL/A16;" '13
        Campos$ = Campos$ + "CINTILLA CARG SEG/A16;" '14
        Campos$ = Campos$ + "CINTILLA LINER/A13;" '15
        Campos$ = Campos$ + "MAMPARO/A11;" '16
        Campos$ = Campos$ + "HILO/A11;" '17
        Campos$ = Campos$ + "ESLINGA/A11;" '18
        Campos$ = Campos$ + "ESLINGA PERIM/A13;" '19
        Campos$ = Campos$ + "ETIQUETA/A11;" '20
        Campos$ = Campos$ + "SOBRE/A11;" '21
        Campos$ = Campos$ + "LINER/A11;" '22
        Campos$ = Campos$ + "ENVOLTORIO/A11;" '23
        Campos$ = Campos$ + "ID/A0" '24

'        CAMPOS$ = "CODIGO/A18;DESCRIPCION/A84;VC/A6;VCS/A6;VD/A6;POLLERA/A8;COMPART/A8;FDO_PANEL/A10;CPO.LAT./A10;"
'        CAMPOS$ = CAMPOS$ + " VALV.CRGA./A10;VALV.DESC./A10;T.POLLERA/A10;CINT.DSC./A10;CINT.DSC.SEG./A12;CINT.CARG.POLL./A14;"
'        CAMPOS$ = CAMPOS$ + " CINT.CARG.SEG./A14;CINT.LINER/A12;MAMPARO/A12;HILO1/A10;HILO2/A10;HILO3/A10;LATERAL2/A10;ESLINGA/A10;"
'        CAMPOS$ = CAMPOS$ + " ESL.PERIM./A12;ETIQUETA/A10;SOBRE/A10;TINTA/A10;DILUYENTE/A10;LINER/A10;ENVOLTORIO/A10;ID/A0"
        VENTANA.Campox = Campos$
        Dim obj As New RECORXET
        Set RS1 = obj.RS_CONFIGBIGBAG()
        Call Carga_MSF_Recordset(RS1, Campos$, VENTABNEW.MSF)
        Set obj = Nothing
        RS1.Close
        Set RS1 = Nothing
        Call numerargrilla(VENTABNEW.MSF, 0, 0, 0, 0)

        VENTABNEW.Show 1
        If VENTABNEW.APROBADO% < 1 Then Exit Sub


        'TRANSACCIÓN GRABACION
        On Error GoTo ERROR_GUARDAR
        FLEXCONN.BeginTrans

        CONDI$ = " CONFIGBB_ID > 0"
        Call ACTUREGISTRO("CONFIGBB", "MARBORRA", CONDI$, 1, REGAF&, "NOMESS")
        
        SQLX$ = "SELECT * FROM CONFIGBB"
        SQLX$ = SQLX$ + " WHERE CONFIGBB_ID < 1"
        Dim MSF As msFlexGrid
        Set MSF = VENTABNEW.MSF
        
        Dim RS As ADODB.Recordset
        Set RS = New ADODB.Recordset
        RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText

          With RS
              For REG& = 1 To MSF.Rows - 1
               .AddNew
                  
                  !CODBIGBAG = MSF.TextMatrix(REG&, 1)
                  !DESCRIPCIONBB = MSF.TextMatrix(REG&, 2)
                  
                  !CUERPOLAT = MSF.TextMatrix(REG&, 3)
                  !LATERAL2 = MSF.TextMatrix(REG&, 4)
                  !FONDOPANEL = MSF.TextMatrix(REG&, 5)
                  !TAPAPOLLERA = MSF.TextMatrix(REG&, 6)
                  !Pollera = MSF.TextMatrix(REG&, 7)
                  !UPANEL = MSF.TextMatrix(REG&, 8)
                  !VALVCARGA = MSF.TextMatrix(REG&, 9)
                  !VALVDESCAR = MSF.TextMatrix(REG&, 10)
                  
                  !CINTIPOLLERA = MSF.TextMatrix(REG&, 11)
                  !CINTDESC = MSF.TextMatrix(REG&, 12)
                  !CINTDESCSEG = MSF.TextMatrix(REG&, 13)
                  !CINTCARGPOLL = MSF.TextMatrix(REG&, 14)
                  !CINTCARGSEG = MSF.TextMatrix(REG&, 15)
                  
                  !CINTLINER = MSF.TextMatrix(REG&, 16)
                  !Mamparo = MSF.TextMatrix(REG&, 17)
                  !HILO1 = MSF.TextMatrix(REG&, 18)
                  !Eslinga = MSF.TextMatrix(REG&, 19)
                  !ESLPERIM = MSF.TextMatrix(REG&, 20)
                  !ETIQUETA = MSF.TextMatrix(REG&, 21)
                  !SOBRE = MSF.TextMatrix(REG&, 22)
                  !LINER = MSF.TextMatrix(REG&, 23)
                  !ENVOLTORIO = MSF.TextMatrix(REG&, 24)
                  !CODIEMPR = CodiEmp%
                  !MARBORRA = 0
                  .Update
              Next REG&
          End With
        
          RS.Close
          Set RS = Nothing
        
        CONDI$ = CONDI$ + " AND  MARBORRA = 1  "
        Call BORRAREGISTROS("CONFIGBB", CONDI$, REGAF&, "NOMESS")
        
        Set VENTANA = Nothing
        
ERROR_GUARDAR:
    If Err <> 0 Then
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

Private Sub mnuCopiaSoloDatosTecnicos_Click()
   If XVALO(Text5) > 0 Then
      If VALIDASISEMODIFICO% < 0 Then Exit Sub
      If CantRegistros&("RECETACTP", "NUMRECETACTP = " & XVALO(Text5)) < 1 Then
         Call COMUNI("Debe Seleccionar Una Cotización Que esté Correctamente Guardada en al Base de Datos")
         Exit Sub
      End If
      
      OPX% = COMALTER%("Confirmación de Grabación de Nueva Ficha Técnica en Base a la Ficha Técnica Activa\\Cancelar\Si Grabar")
      If OPX% > 1 Then
         LBLNUEVAFICHATECNICA.Visible = True
         TXTFECHA = FECHATEX$(HOY(HO$))
         Text1 = ""
         TXTRAZOCLI = ""
         TXTOBSERVA = ""
         TXTCODESPECIFTECCLTE = ""
         TXTCODIGO = ""
         TXTNUMCOTIZACION = ""
         TXTDESCRICOD = ""
         TXTDESCRICODCONFECCION = ""
         Text2 = ""
         TXTJUSTA = "NO"
         TXTCAPACIDAD = ""
         TXTPRODUCTO = ""
         TXTCODCONFECCION = ""
         TXTCANTIUSOS = ""
         
         Call MnuGuardar_Click
         
         LBLNUEVAFICHATECNICA.Visible = False
         MODIFIC% = 0
      End If
   End If

End Sub

Private Sub mnuCodiRang_Click()
  EPAC$ = TRIM$(UCase$(EMPREAC$))
  If InStr(1, EPAC$, "CONTEPLAST") > 0 Then
    If DERACCE%("SOYVENDE", "Derecho Bloqueado") < 2 Then Exit Sub
       'sale solo por que al perder el foco no entra en el menu
  End If

    Erase RESP_ZELE_VECT
    '29/09/2020
    Campos$ = "Cód.Rango/A10;" '1
    Campos$ = Campos$ + "Descripción/A64;"      '2
    Campos$ = Campos$ + "Cód.M.P./A16;"           '3
    Campos$ = Campos$ + "DESCRIPCION/A28;"      '4
    Campos$ = Campos$ + "Costo/F10.4;"      '5
    Campos$ = Campos$ + "CIMAT/A0"      '6
    
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    FRMZELECHO.Caption = ""
    SQLX$ = "SELECT * FROM CODIGORANGO "
    SQLX$ = SQLX$ + " WHERE CODIGORANGO_ID < 1"
    Dim obj As New RECORXET
    CAMPOSTABLA$ = "C.CODRANGO ,C.Descripcion ,C.CODMATPRIMA ,C.DESCRIMATPRIMA,M.COSUNI,C.CIMATPRIMA"
    TABLA$ = "CODIGORANGO C WITH(NOLOCK) INNER JOIN MASTER M ON C.CIMATPRIMA = M.CODINT"
    CONDI$ = "C.CODIGORANGO_ID > 0 Order by C.CODRANGO"
    '
    Set RS1 = obj.RS_GENERAL(TABLA$, CAMPOSTABLA$, CONDI$)
    Call Carga_MSF_Recordset(RS1, Campos$, FRMZELECHO.msf2, 1)
    On Error Resume Next
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    On Error GoTo 0
        
    FRMZELECHO.Width = 15705
50  FRMZELECHO.Show 1

End Sub

Private Sub mnuCodRango_Click()
        If DERACCE("CONCODRA", "CONFIGURACION DE CODIGOS RANGOS") < 2 Then Exit Sub
        Unload VENTABNEW
        VENTANA.Inicializar = 0
        VENTANA.CampEditables = "1;3"
        VENTANA.AgregaRegistro = 1
        VENTANA.UtilizaFormula = 1
        VENTANA.HabilitaBorrar = 1
        VENTANA.HabilitaInsertar = 1
        VENTANA.NoMostrarBotonTilde = 0
        VENTANA.TITULO = "CONFIGURACIÓN CODIGOS RANGO"
        VENTANA.ColEcoCampo(1) = "CODIGO;DESCRIPCION;MASTER;Timate  = 49"
        VENTANA.ColEcoCampo(3) = "CODIGO;DESCRIPCION;MASTER;Timate  <> 49"
        VENTANA.Inicializar = 1
        Campos$ = "Cód.Rango/A16;" '1
        Campos$ = Campos$ + "Descripción/A64;"      '2
        Campos$ = Campos$ + "Cód.M.P./A16;"           '3
        Campos$ = Campos$ + "DESCRIPCION/A64;"      '4
        Campos$ = Campos$ + "CIR/A0;"       '5
        Campos$ = Campos$ + "CIMP/A0;"       '6
        Campos$ = Campos$ + "ID/A0" '7
        VENTANA.Campox = Campos$
        Dim obj As New RECORXET
        Set RS1 = obj.RS_CODRANGO()
        Call Carga_MSF_Recordset(RS1, Campos$, VENTABNEW.MSF)
        Set obj = Nothing
        RS1.Close
        Set RS1 = Nothing
        Call numerargrilla(VENTABNEW.MSF, 0, 0, 0, 0)
'
20      VENTABNEW.Show 1
        If VENTABNEW.APROBADO% < 1 Then Exit Sub


        'TRANSACCIÓN GRABACION
        On Error GoTo ERROR_GUARDAR
        FLEXCONN.BeginTrans
        CONDI$ = "CODIGORANGO_ID > 0"
        Call ACTUREGISTRO("CODIGORANGO", "MARBORRA", CONDI$, 1, REGAF&, "NOMESS")
        SQLX$ = "SELECT * FROM CODIGORANGO"
        SQLX$ = SQLX$ + " WHERE CODIGORANGO_ID < 1"
        Dim MSF As msFlexGrid
        Set MSF = VENTABNEW.MSF
        
        Dim RS As ADODB.Recordset
        Set RS = New ADODB.Recordset
        RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText

        With RS
              For REG& = 1 To MSF.Rows - 1
                  If TRIM$(MSF.TextMatrix(REG&, 1)) = "" Then GoTo 50
                  .AddNew
                  !CODRANGO = MSF.TextMatrix(REG&, 1)
                  !Descripcion = MSF.TextMatrix(REG&, 2)
                  !CODMATPRIMA = MSF.TextMatrix(REG&, 3)
                  !DESCRIMATPRIMA = MSF.TextMatrix(REG&, 4)
                  CIRGO = CODINT%(MSF.TextMatrix(REG&, 1))
                  !CIRANGO = CIRGO
                  MSF.TextMatrix(REG&, 5) = CIRGO ' SI NO LO TENIA LO ESCRIBE EN LA GRILLA PARA VALIDAR ABAJO
                  CIMP = CODINT%(MSF.TextMatrix(REG&, 3))
                  !CIMATPRIMA = CIMP
                  MSF.TextMatrix(REG&, 6) = CIMP ' SI NO LO TENIA LO ESCRIBE EN LA GRILLA PARA VALIDAR ABAJO
                  !MARBORRA = 0
                  .Update
50            Next REG&
        End With
        
        RS.Close
        Set RS = Nothing
        
        '
        'VALIDA AQUI SI COINCIDEN LAS U.MEDIDA (LO HAGO AQUI PARA NO REPETIR ANTES ASIGNAR EL COD INTERNO.)
        For REG& = 1 To MSF.Rows - 1
           CIRANGOX% = XVALO(MSF.TextMatrix(REG&, 5))
           CIMPX% = XVALO(MSF.TextMatrix(REG&, 6))
           If CIRANGOX% >= 0 And CIMPX% >= 0 Then
               If Unimed$(CIRANGOX%) <> Unimed$(CIMPX%) Then
                 Call COMUNI("Atención En Fila: " & REG& & " No Coincide Tipo de Material\" & MSF.TextMatrix(REG&, 1) & " = " & Unimed$(CIRANGOX%) & "\" & MSF.TextMatrix(REG&, 3) & " = " & Unimed$(CIMPX%))
                 FLEXCONN.RollbackTrans
                 On Error GoTo 0
                 GoTo 20
               End If
           End If
        Next REG&
        
        'VALIDO QUE EL CODIGO DE MATERIA PRIMA NO ESTE REPETIDO
        For REG& = 1 To MSF.Rows - 1
           CIRANGOX% = XVALO(MSF.TextMatrix(REG&, 5))
           CIMPX% = XVALO(MSF.TextMatrix(REG&, 6))
           For REG2& = 1 To MSF.Rows - 1
             CIMPX2% = XVALO(MSF.TextMatrix(REG2&, 6))
             If REG2& <> REG& And CIMPX2% = CIMPX% Then
               Call COMUNI("Atención En Fila: " & REG& & " Código: " & CODEXT$(CIMPX%) & " Repetido en Fila: " & REG2&)
               FLEXCONN.RollbackTrans
               On Error GoTo 0
               GoTo 20
             End If
           Next REG2&
        Next REG&
        
        CONDI$ = CONDI$ + " AND  MARBORRA = 1  "
        Call BORRAREGISTROS("CODIGORANGO", CONDI$, REGAF&, "NOMESS")
        
        Set VENTANA = Nothing
        
ERROR_GUARDAR:
    If Err <> 0 Then
90     FLEXCONN.RollbackTrans
       On Error GoTo 0
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION")
       
    Else
      FLEXCONN.CommitTrans
      Call COMUNI("Grabación Realizado Exitosamente")
    End If
    Unload VENTABNEW
    
   
End Sub

Private Sub mnuConfigWaldbot_Click()
   USER1% = USNBR% Mod 100
   If USER1% <> 1 Then
      Call COMUNI("Solo Puede Ingresar el Supervisor del Sistema")
      Exit Sub
   End If

   BkpWaldb.Show 1
   
End Sub

Private Sub mnucortebigbag_Click()
        If DERACCE("CONCALCO", "CONFIGURACION DE CALCULO DE CORTE") < 2 Then Exit Sub
        Unload VENTABNEW
        VENTANA.Inicializar = 0
        VENTANA.CampEditables = "1;3;4;5;6;7;8"
        VENTANA.AgregaRegistro = 1
        VENTANA.UtilizaFormula = 1
        VENTANA.HabilitaBorrar = 1
        VENTANA.HabilitaInsertar = 1
        VENTANA.NoMostrarBotonTilde = 0 'A*2+L*2+7
        VENTANA.TITULO = "FORMULAS CORTE (TEXTO LIMITADOS x (;) EJ:ALTO;*2+;LARGO;*2+7)"
        VENTANA.ColEcoCampo(1) = "CAMPO1;CAMPO2;DATASQL;LISTA  = 'CODBB'"
        VENTANA.Inicializar = 1
        Campos$ = "CODIGO/A18;" '0
        Campos$ = Campos$ + "DESCRIPCION/A84;"      '1
        Campos$ = Campos$ + "CUERPO C.C./A24;"      '2
        Campos$ = Campos$ + "LATERALES C.C./A24;"   '3
        Campos$ = Campos$ + "FONDO C.C./A24;"       '4
        Campos$ = Campos$ + "TAPA C.C./A24;"        '5
        Campos$ = Campos$ + "POLLERA C.C./A24;"     '6
        Campos$ = Campos$ + "U PANEL C.C./A24;"     '7
        Campos$ = Campos$ + "ID/A0" '8

        VENTANA.Campox = Campos$
        
        Dim obj As New RECORXET

        CAMPOSTABLA$ = "CODBIGBAG,DESCRIPCIONBB,CUERPOLAT,LATERAL2,FONDOPANEL,TAPAPOLLERA,POLLERA,UPANEL"
        Set RS1 = obj.RS_GENERAL("CALCUCORTE", CAMPOSTABLA$, "CALCUCORTE_ID > 0")
        Call Carga_MSF_Recordset(RS1, Campos$, VENTABNEW.MSF)
        On Error Resume Next
        Set obj = Nothing
        RS1.Close
        Set RS1 = Nothing
        Call numerargrilla(VENTABNEW.MSF, 0, 0, 0, 0)
        On Error GoTo 0
        
        VENTABNEW.Show 1
        If VENTABNEW.APROBADO% < 1 Then Exit Sub


        'TRANSACCIÓN GRABACION
        On Error GoTo ERROR_GUARDAR
        FLEXCONN.BeginTrans

        CONDI$ = " CALCUCORTE_ID > 0"
        Call ACTUREGISTRO("CALCUCORTE", "MARBORRA", CONDI$, 1, REGAF&, "NOMESS")
        
        SQLX$ = "SELECT * FROM CALCUCORTE"
        SQLX$ = SQLX$ + " WHERE CALCUCORTE_ID < 1"
        Dim MSF As msFlexGrid
        Set MSF = VENTABNEW.MSF
        
        Dim RS As ADODB.Recordset
        Set RS = New ADODB.Recordset
25      On Error Resume Next
        RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        If Err Then
         On Error GoTo 0
         Call CapturaErrorOpen
         GoTo 25
        End If
          With RS
              For REG& = 1 To MSF.Rows - 1
               .AddNew
                  !CODBIGBAG = MSF.TextMatrix(REG&, 1)
                  !DESCRIPCIONBB = MSF.TextMatrix(REG&, 2)
                  !CUERPOLAT = MSF.TextMatrix(REG&, 3)
                  !LATERAL2 = MSF.TextMatrix(REG&, 4)
                  !FONDOPANEL = MSF.TextMatrix(REG&, 5)
                  !TAPAPOLLERA = MSF.TextMatrix(REG&, 6)
                  !Pollera = MSF.TextMatrix(REG&, 7)
                  !UPANEL = MSF.TextMatrix(REG&, 8)
                  !CODIEMPR = CodiEmp%
                  .Update
              Next REG&
          End With
        
          RS.Close
          Set RS = Nothing
        
        CONDI$ = CONDI$ + " AND  MARBORRA = 1  "
        Call BORRAREGISTROS("CALCUCORTE", CONDI$, REGAF&, "NOMESS")
        Set VENTANA = Nothing
        
ERROR_GUARDAR:
    If Err <> 0 Then
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

Private Sub mnuCostos_Click()
''''''
   If DERACCE%("REPCOSTO", "REPORTE DE COSTOS") < 2 Then Exit Sub
   Screen.MousePointer = 11

20 Call DESHASFECHA(DESDEX%, HASTAX%, PERIO$)
   If PERIO$ = "" Then Exit Sub

   DESFE = FETEXCOR1$(DESDEX%)
   HASFE = FETEXCOR1$(HASTAX%, "HASTA")

   Campos$ = "FECHA/D8" '1
   Campos$ = Campos$ + ";OFAB/a12" '2
   Campos$ = Campos$ + ";PROD/F9.3" '3
   Campos$ = Campos$ + ";COD CONF/A16" '4
   Campos$ = Campos$ + ";CANT/F8" '5
   Campos$ = Campos$ + ";TELA/F9.4" '6
   Campos$ = Campos$ + ";LINER/F9.4" '7
   Campos$ = Campos$ + ";OTROS/F9.4" '8
   Campos$ = Campos$ + ";SCRAP/F9.4" '9

   Campos$ = Campos$ + ";COMISION/F9.4" '10
   Campos$ = Campos$ + ";MO/F9.4" '11
   Campos$ = Campos$ + ";COSTO TOT/F9.4" '12
   Campos$ = Campos$ + ";MARGEN/F9.4" '13
   Campos$ = Campos$ + ";N PED/f8" '14
   Campos$ = Campos$ + ";P VENTA/F9.4" '15
   Campos$ = Campos$ + ";NªCotiz./F9" '16

   AcuCanti = 0
   AcuCOSTOTELAS# = 0
   AcuLINER = 0
   AcuOTROSCOSTOS# = 0
   AcuSCRAP = 0
   AcuMARGEN = 0
   AcuCOSTOTOTAL = 0
   AcuCOSTOMANOOBRA = 0
   
   Dim CTP As New Conteplast


   Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
   Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)
   J& = 0
   SQLX$ = " SELECT MOV.FECHMOV,MOV.DOPRICOR,MOV.COD_INTE,MOV.COD_EXTE,SUM(MOV.CANTINGRE) as CANTINGRE FROM MOVISTO MOV WITH(NOLOCK) "
   SQLX$ = SQLX$ + " INNER JOIN MASTER MAS ON MOV.COD_INTE = MAS.CODINT "
   SQLX$ = SQLX$ + " WHERE MOV.CODIMOVI = 'OF' AND MOV.CANTINGRE > 0 "
   SQLX$ = SQLX$ + " AND MOV.FECHMOV >= '" & DESFE & "'"
   SQLX$ = SQLX$ + " AND MOV.FECHMOV <= '" & HASFE & "'"
   SQLX$ = SQLX$ + " AND MAS.FAMILIA = 'BB'"

   SQLX$ = SQLX$ + " GROUP BY FECHMOV,DOPRICOR,COD_INTE,COD_EXTE"
   SQLX$ = SQLX$ + " ORDER BY FECHMOV,DOPRICOR,COD_INTE,COD_EXTE"
   Set RS = READSET(SQLX$)
   With RS
      Do While Not .EOF
       J& = J& + 1
       FE% = CVINT(!FechMov)
       NORFA$ = SAFETEXT$(!dopricor)
       CI1% = XVALO(!cod_inte)
       Codigo$ = SAFETEXT$(!Cod_Exte)
       COBIBAG$ = CTP.CODIGOBIGBAG$(CI1%)
       CANTI = XVALO(!CANTINGRE)
       NPED& = NroPedidoSegunOf&(NORFA$)
       NroOrdenPedido% = NumeroOrdenPedidoSegunOrdenFabricacion%(NORFA$)
       PREUNIDAD# = PrecioUnitarioSegunPedido#(NPED&, CI1%, NroOrdenPedido%)
       NROCOTIZAX = NROCOTIZA(NPED&, NroOrdenPedido%)
       LINER = CTP.CostoLinerSegunCotizacion2#(NROCOTIZAX)
       COSTOTELAS# = CTP.CostoTelasSegunCotizacion1#(NROCOTIZAX)
       OTROSCOSTOS# = CTP.SUMADEOTROSSegunCotizacion2#(NROCOTIZAX)
       Scrap = CTP.ScrapSegunCotizacion2#(NROCOTIZAX)
       COMISION = CTP.ComisionSegunCotizacion2#(NROCOTIZAX)
       COSTOMANOOBRA = CTP.ManoObraSegunCotizacion2#(NROCOTIZAX)
       COSTOTOTAL = CTP.CostoTotalSegunCotizacion2#(NROCOTIZAX)
       CostoTotalyComisionYLiner = COSTOTOTAL + COMISION + LINER
       PJEUTILIDAD = CTP.PorcentajeUtilidadSegunCotizacion2#(NROCOTIZAX)
       MARGEN = COSTOTOTAL * PJEUTILIDAD / 100
           GRILLAGRAL.GRID.Rows = J& + 1
           GRILLAGRAL.GRID.TextMatrix(J&, 1) = FECHATEX$(FE%)
           GRILLAGRAL.GRID.TextMatrix(J&, 2) = NORFA$
           GRILLAGRAL.GRID.TextMatrix(J&, 3) = Codigo$
    
           GRILLAGRAL.GRID.TextMatrix(J&, 4) = COBIBAG$
           GRILLAGRAL.GRID.TextMatrix(J&, 5) = CANTI: AcuCanti = AcuCanti + CANTI
           GRILLAGRAL.GRID.TextMatrix(J&, 6) = TRIM$(FORMATNUM$(COSTOTELAS#, "F9.4")): AcuCOSTOTELAS# = AcuCOSTOTELAS# + COSTOTELAS#
    
           GRILLAGRAL.GRID.TextMatrix(J&, 7) = TRIM$(FORMATNUM$(LINER, "F9.4")): AcuLINER = AcuLINER + LINER
           GRILLAGRAL.GRID.TextMatrix(J&, 8) = TRIM$(FORMATNUM$(OTROSCOSTOS#, "F9.4")): AcuOTROSCOSTOS# = AcuOTROSCOSTOS# + OTROSCOSTOS#
           GRILLAGRAL.GRID.TextMatrix(J&, 9) = TRIM$(FORMATNUM$(Scrap, "F9.4")): AcuSCRAP = AcuSCRAP + Scrap
           
           GRILLAGRAL.GRID.TextMatrix(J&, 10) = TRIM$(FORMATNUM$(COMISION, "F9.4")):
           GRILLAGRAL.GRID.TextMatrix(J&, 11) = TRIM$(FORMATNUM$(COSTOMANOOBRA, "F9.4")): AcuCOSTOMANOOBRA = AcuCOSTOMANOOBRA + COSTOMANOOBRA
           GRILLAGRAL.GRID.TextMatrix(J&, 12) = TRIM$(FORMATNUM$(CostoTotalyComisionYLiner, "F9.4")): AcuCOSTOTOTAL = AcuCOSTOTOTAL + CostoTotalyComisionYLiner
           GRILLAGRAL.GRID.TextMatrix(J&, 13) = TRIM$(FORMATNUM$(MARGEN, "F9.4")): AcuMARGEN = AcuMARGEN + MARGEN
           GRILLAGRAL.GRID.TextMatrix(J&, 14) = TRIM$(FORMATNUM$(NPED&, "F7"))
           GRILLAGRAL.GRID.TextMatrix(J&, 15) = TRIM$(FORMATNUM$(PREUNIDAD#, "F9.3"))
           GRILLAGRAL.GRID.TextMatrix(J&, 16) = CStr(NROCOTIZAX)
           .MoveNext
     Loop
    End With
    RS.Close
    Set RS = Nothing
    '
    Screen.MousePointer = SCMA
    
     TX$ = "Reporte de Analisis de Costos (" & PERIO$ & ")"
     GRILLAGRAL.Caption = TX$
     '*******************************************
     GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
     GRILLAGRAL.mnuMenu1.Visible = True
     GRILLAGRAL.GRID.ZOrder 0
     GRILLAGRAL.TXTBUSCAR = ""
     Call POSICIONAR_ULTIMAFILA(GRILLAGRAL.GRID)
     
     GRILLAGRAL.SinEfectoClickEnGrilla% = 1
     
     
     
     GRILLAGRAL.Show 1
     Dim MSF As msFlexGrid
     If GRILLAGRAL.MENU1 > 0 Then
         Set MSF = GRILLAGRAL.GRID
         If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
         Call CARGA_TABLA_IMPRE(MSF, Campos$, "REPOANCO", TX$, "TABLA_IMPRE", RET_NAMECONS$, 1)
         Call FINAL("?" & RET_NAMECONS$)
         GRILLAGRAL.MENU1 = 0
         Screen.MousePointer = 1
     End If

End Sub

Private Sub mnuCotizaciones_Click()
  EPAC$ = TRIM$(UCase$(EMPREAC$))
  If InStr(1, EPAC$, "CONTEPLAST") > 0 Then
    If DERACCE%("SOYVENDE", "Derecho Bloqueado") < 2 Then Exit Sub
       'sale solo por que al perder el foco no entra en el menu
  End If

   Screen.MousePointer = 11
   Campos$ = "Cotiz./F8;F.Téc./F8;Cta./F8;R.Social/A24;Cód.Conf./A8;Código/A16;F.Emis/D8;Pedido/f8;Fech.Ped./D8;Vend./F6"
   
   Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
   Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)
   J& = 0
    SQLX$ = " select  c.NROCOTIZACION,c.NUMRECETACTP as FT ,C.NUCLIEN,C.RASOCLI ,"
    SQLX$ = SQLX$ + " S.CODBIGBAG, c.CODIMASTER , C.FECHEMISI AS FECOT,P.NroPed ,P.FechEmis AS FEPED,P.Vended "
    SQLX$ = SQLX$ + " ,C.STATUS ,O.IdSubOr "
    SQLX$ = SQLX$ + " from cotizactp c with(nolock) "
    SQLX$ = SQLX$ + " LEFT JOIN SOLICOTI S ON S.CODIMASTER = C.CODIMASTER"
    SQLX$ = SQLX$ + " LEFT JOIN peddeta p on c.CODIMASTER = p.CodiExt  and c.NROCOTIZACION = p.NROCOTI "
    SQLX$ = SQLX$ + " LEFT JOIN ORDEFABR O  ON P.CodiInt = O.Cod_Inte AND O.NroPed = P.NroPed "
    SQLX$ = SQLX$ + " ORDER BY C.NROCOTIZACION DESC "
   Set RS = READSET(SQLX$)
   With RS
     Do While Not .EOF
       J& = J& + 1
       '
       NRO& = XVALO(!NROCOTIZACION)
       FT& = XVALO(!FT)
       NC = XVALO(!NUCLIEN)
       RACLI$ = SAFETEXT$(!rasocli)
       CODIBB$ = SAFETEXT$(!CODBIGBAG)
       CODIMAS$ = SAFETEXT$(!CODIMASTER)
       FECHAEMI$ = FECHATEX$(CVINT(!FECOT))
       NROPED& = XVALO(!NROPED)
       FEPED$ = FECHATEX$(CVINT(!FEPED))
       NVEN = XVALO(!VENDED)
       GRILLAGRAL.GRID.Rows = J& + 1
       GRILLAGRAL.GRID.TextMatrix(J&, 1) = NRO&
       GRILLAGRAL.GRID.TextMatrix(J&, 2) = FT&
       GRILLAGRAL.GRID.TextMatrix(J&, 3) = NC
       
       GRILLAGRAL.GRID.TextMatrix(J&, 4) = RACLI$
       GRILLAGRAL.GRID.TextMatrix(J&, 5) = CODIBB$
       GRILLAGRAL.GRID.TextMatrix(J&, 6) = CODIMAS$
        
       GRILLAGRAL.GRID.TextMatrix(J&, 7) = FECHAEMI$
       GRILLAGRAL.GRID.TextMatrix(J&, 8) = NROPED&
       GRILLAGRAL.GRID.TextMatrix(J&, 9) = FEPED$
       GRILLAGRAL.GRID.TextMatrix(J&, 10) = NVEN
       .MoveNext
     Loop
   End With
   RS.Close
   Set RS = Nothing
   '
   Screen.MousePointer = SCMA
    
    TX$ = "Cotizaciones"
    GRILLAGRAL.Caption = TX$
    '*******************************************
    GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
    GRILLAGRAL.mnuMenu1.Visible = True
    GRILLAGRAL.GRID.ZOrder 0
    GRILLAGRAL.TXTBUSCAR = ""
    Call POSICIONAR_ULTIMAFILA(GRILLAGRAL.GRID)
    
    GRILLAGRAL.SinEfectoClickEnGrilla% = 1
    
'    If XCODIGO% > 0 Then
'      GRILLAGRAL.CMBFILTRO = "Código"
'      GRILLAGRAL.TXTBUSCAR = CODEXT$(XCODIGO%)
'    Else
'      GRILLAGRAL.CMBFILTRO = "Cta."
'      NumeroCliente$ = CStr(XCLIENTE)
'      While Len(NumeroCliente$) < 7: NumeroCliente$ = "0" & NumeroCliente$: Wend
'      GRILLAGRAL.TXTBUSCAR = NumeroCliente$
'    End If
    
    GRILLAGRAL.Show 1
    Dim MSF As msFlexGrid
    If GRILLAGRAL.MENU1 > 0 Then
        Set MSF = GRILLAGRAL.GRID
        If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
        Call CARGA_TABLA_IMPRE(MSF, Campos$, "ENCAPEDI", TX$, "TABLA_IMPRE", RET_NAMECONS$, 1)
        Call FINAL("?" & RET_NAMECONS$)
        GRILLAGRAL.MENU1 = 0
        Screen.MousePointer = 1
    End If

End Sub

Private Sub mnuftcamba_Click()
End Sub

Private Sub mnuftecnicas_Click()
  EPAC$ = TRIM$(UCase$(EMPREAC$))
  If InStr(1, EPAC$, "CONTEPLAST") > 0 Then
    If DERACCE%("SOYVENDE", "Derecho Bloqueado") < 2 Then Exit Sub
       'sale solo por que al perder el foco no entra en el menu
  End If

   Screen.MousePointer = 11
   Campos$ = "F.Téc./F8;Emisión/D10;Cta./F8;R.Social/A24;Cód.Conf./A8;Código/A16;Descripción/A64;F.Aprob/D8.1;Vend./F6;Ci/a0"
   
   Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
   Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)
   J& = 0
   SQLX$ = " SELECT  "
   SQLX$ = SQLX$ + " NUMRECETACTP,FECHEMISI,NUCLIEN,RASOCLI, CODBIGBAG,CODIMASTER,M.DESCRIPCION,FEAPROBO,R.NUVEND,R.CODINT "
   SQLX$ = SQLX$ + " from RECETACTP R WITH(NOLOCK) INNER JOIN MASTER M ON R.CODINT = M.CODINT"
   SQLX$ = SQLX$ + " ORDER BY FECHEMISI DESC"
   
   Set RS = READSET(SQLX$)
   With RS
     Do While Not .EOF
       J& = J& + 1
       '
       NRO& = XVALO(!NUMRECETACTP)
       FechEmiX$ = FECHATEX$(CVINT(!FECHEMISI))
       NC = XVALO(!NUCLIEN)
       RACLI$ = SAFETEXT$(!rasocli)
       CODIBB$ = SAFETEXT$(!CODBIGBAG)
       CODIMAS$ = SAFETEXT$(!CODIMASTER)
       DESKRIP$ = SAFETEXT$(!Descripcion)
       FECHAPRO$ = FECHATEX$(CVINT(!FEAPROBO))
       NVEN = XVALO(!NUVEND)
       CI1% = XVALO(!CODINT)
       GRILLAGRAL.GRID.Rows = J& + 1
       GRILLAGRAL.GRID.TextMatrix(J&, 1) = NRO&
       GRILLAGRAL.GRID.TextMatrix(J&, 2) = FechEmiX$
       GRILLAGRAL.GRID.TextMatrix(J&, 3) = NC
       
       GRILLAGRAL.GRID.TextMatrix(J&, 4) = RACLI$
       GRILLAGRAL.GRID.TextMatrix(J&, 5) = CODIBB$
       GRILLAGRAL.GRID.TextMatrix(J&, 6) = CODIMAS$
        
       GRILLAGRAL.GRID.TextMatrix(J&, 7) = DESKRIP$
       GRILLAGRAL.GRID.TextMatrix(J&, 8) = FECHAPRO$
       GRILLAGRAL.GRID.TextMatrix(J&, 9) = NVEN
       GRILLAGRAL.GRID.TextMatrix(J&, 10) = CI1%
       .MoveNext
     Loop
   End With
   RS.Close
   Set RS = Nothing
   '
   Screen.MousePointer = SCMA
    
    TX$ = "F.Técnicas"
    GRILLAGRAL.Caption = TX$
    '*******************************************
    GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
    GRILLAGRAL.mnuMenu1.Visible = True
    GRILLAGRAL.GRID.ZOrder 0
    GRILLAGRAL.TXTBUSCAR = ""
    Call POSICIONAR_ULTIMAFILA(GRILLAGRAL.GRID)
    
    GRILLAGRAL.SinEfectoClickEnGrilla% = 1
    
'    If XCODIGO% > 0 Then
'      GRILLAGRAL.CMBFILTRO = "Código"
'      GRILLAGRAL.TXTBUSCAR = CODEXT$(XCODIGO%)
'    Else
'      GRILLAGRAL.CMBFILTRO = "Cta."
'      NumeroCliente$ = CStr(XCLIENTE)
'      While Len(NumeroCliente$) < 7: NumeroCliente$ = "0" & NumeroCliente$: Wend
'      GRILLAGRAL.TXTBUSCAR = NumeroCliente$
'    End If
    
    GRILLAGRAL.Show 1
    Dim MSF As msFlexGrid
    If GRILLAGRAL.MENU1 > 0 Then
        Set MSF = GRILLAGRAL.GRID
        If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
        Call CARGA_TABLA_IMPRE(MSF, Campos$, "ENCAPEDI", TX$, "TABLA_IMPRE", RET_NAMECONS$, 1)
        Call FINAL("?" & RET_NAMECONS$)
        GRILLAGRAL.MENU1 = 0
        Screen.MousePointer = 1
    End If

End Sub

Private Sub mnuGuardarCoti_Click()
     Call Command46_Click
     Call FRMCOTI2.Command47_Click
End Sub

Private Sub mnuimprimecotizacion_Click()

    If DERACCE%("IMPRCOTI", "IMPRIMIR COTIZACIONES") < 2 Then Exit Sub
    Dim CTP As New Conteplast
    CONDI$ = " AND (FT_STAT < 9 OR FT_STAT IS NULL) AND COT_NROCOT > 0 "
    
    NROFT = CTP.SELEC_FTECNICAYCOTI(CONDI$)
    If NROFT < 1 Then Exit Sub
    
    NROCOT& = XVALO(RESP_ZELE_VECT(3))
    NroVend% = XVALO(RESP_ZELE_VECT(11))
    FE% = HOY(HO$)
    VENDEDOR$ = TRIM$(ECOARCH$("VENDEDOR", Chr$(NroVend%)))
    FECHALARGA$ = FormatDateTime(FECHATEX$(FE%), vbLongDate)
    'para que en vez de lunes, 3 de febrero quede Lunes, 3 de Febrero (mayuscula en dia y mes)
    Dim S() As String, i As Integer
    S = Split(FECHALARGA$, " ")
    For i = LBound(S) To UBound(S)
       If S(i) <> "de" Then
         Mid$(S(i), 1, 1) = UCase$(Mid$(S(i), 1, 1))
       End If
    Next
    FECHALARGA$ = Join$(S, " ") ' concatena todo el vector

    
    Call SETPRINPORT("RESTORE")
    PUERTOIMPRE$ = Control$(IDENTER$, "PRINTSER")
    If PUERTOIMPRE$ = "" Then PUERTOIMPRE$ = Control$("*", "PRINTSER")
    If PUERTOIMPRE$ <> "" Then
       Call SETPRINPORT(PUERTOIMPRE$)
    End If
    '
    OPX% = COMALTER%("Selección Empresa\\Cancelar\Conteplast\Envaplast")
    
    FILTX$ = FECHALARGA$ & ";" & SAFETEXT$(NROFT) & ";" & SAFETEXT$(NROCOT&) & ";" & ";" & VENDEDOR$ & ";"
    If OPX% = 2 Then
      Call STDFORM("COIN-ENV", FILTX$)
    ElseIf OPX% = 3 Then
      Call STDFORM("COIN-CTP", FILTX$)
    End If

    Call SETPRINPORT("RESTORE")
End Sub

Public Sub mnuimprimirfichaactiva_Click()
    If XVALO(Text5) < 1 Then Exit Sub
    Call PrintFichaTecnica
    Exit Sub
    
    Call SETPRINPORT("RESTORE")
    PUERTOIMPRE$ = Control$(IDENTER$, "PRINTSER")
    If PUERTOIMPRE$ = "" Then PUERTOIMPRE$ = Control$("*", "PRINTSER")
    If PUERTOIMPRE$ <> "" Then
       Call SETPRINPORT(PUERTOIMPRE$)
    End If
    
    NC% = XVALO(Text1)
    
    FILTX$ = TRIM$(Text5) & ";" & TRIM$(Str$(NC%)) & ";" & TRIM$(Str$(XVALO(TXTNUMCOTIZACION)))
    
    OPX% = COMALTER%("Selección Empresa\\Cancelar\Conteplast\Envaplast")
    
    If OPX% = 2 Then
      Call STDFORM("RCTA-CTP", FILTX$)
    ElseIf OPX% = 3 Then
      Call STDFORM("RCT1-ENV", FILTX$)
    End If
    
    Call SETPRINPORT("RESTORE")
End Sub
Sub PrintFichaTecnica()
    Call SETPRINPORT("RESTORE")
    PUERTOIMPRE$ = Control$(IDENTER$, "PRINTSER")
    If PUERTOIMPRE$ = "" Then PUERTOIMPRE$ = Control$("*", "PRINTSER")
    If PUERTOIMPRE$ <> "" Then
       Call SETPRINPORT(PUERTOIMPRE$)
    End If
    '
    RUTAEXCEL$ = Control("COTICONT", "EXCEFTEC")
    If EXISTE(RUTAEXCEL$) < 1 Then
       Call COMUNI("Falta Configurar Ruta de Plantilla Excel de Formulario de Ficha Técnica")
       Exit Sub
    End If
    '
    FTEC& = XVALO(Text5)
    With VENTABNEW.MSF
       .Rows = 68: .Cols = 5
       .TextMatrix(1, 1) = TXTRAZOCLI 'nombre cliente
       .TextMatrix(2, 1) = Text5 'antes TXTCODESPECIFTECCLTE 'Especificacion interna del cliente
       If TRIM$(TXTFECHAPROBO.TEXT) <> "" Then
          .TextMatrix(3, 1) = TXTFECHAPROBO.TEXT
       Else
          .TextMatrix(3, 1) = TXTFECHA.TEXT
       End If
       
       .TextMatrix(4, 1) = "" 'Text5 '
       .TextMatrix(6, 1) = TXTJUSTA 'Cantidad Justa
       
       .TextMatrix(7, 1) = "" 'fecha de recepcion de orden de compra
       .TextMatrix(8, 1) = TXTOBSERVA 'Especificación del Cliente:
       .TextMatrix(9, 1) = "" 'orden de compra
       .TextMatrix(10, 1) = Text3 'Vendedor
       .TextMatrix(11, 1) = TXTCODIGO 'nombre del BB
       .TextMatrix(12, 1) = "" 'Codigo
       
       .TextMatrix(13, 1) = TXTCODCONFECCION 'Codigo de confeccion
       .TextMatrix(14, 1) = TXTDESCRICODCONFECCION 'Descripcion
       .TextMatrix(15, 1) = Text29 'Especificación Impresión del cliente
       .TextMatrix(16, 1) = Text31 'Especificación Impresión Conteplast
       .TextMatrix(17, 1) = TXTCANTICARAS 'Cantidad de Caras
        If TRIM$(Text32) = "" Then Dato$ = Text7 Else Dato$ = Text32
       .TextMatrix(18, 1) = Dato$ 'Colores
       .TextMatrix(19, 1) = TXTTELAVENTILADA 'Tela Ventilada



       .TextMatrix(20, 1) = TXTANCHOCM 'ANCHO
       .TextMatrix(21, 1) = TXTCUERPO_LARGOCM 'LARGO
       .TextMatrix(22, 1) = TXTCUERPOALTOCM 'ALTO
       .TextMatrix(23, 1) = "" 'espacios
       .TextMatrix(24, 1) = "" 'espacios
       .TextMatrix(25, 1) = TXTESLINGAOJOCM 'Eslinga (Ojo)
       .TextMatrix(26, 1) = "" 'espacios
       .TextMatrix(27, 1) = TXTESLINGAMEDIDA1CM 'Eslinga (Medida1)
       .TextMatrix(28, 1) = TXTESLINGAMED2EDITABLE 'Eslinga (Medida2)
       .TextMatrix(29, 1) = "" 'espacios
       .TextMatrix(30, 1) = "" 'espacios
       .TextMatrix(31, 1) = TXTVALVVCDIAMETROCM 'VC (Diámetro)
       .TextMatrix(32, 1) = TXTVCLARGOCM 'VC (LARGO)
       .TextMatrix(33, 1) = "" 'espacios
       .TextMatrix(34, 1) = Text19 'TXTVALVVCLARGOCM 'vcs largo cm se cambia para que imprima el valor
       
       .TextMatrix(35, 1) = "" 'Doble Costura VC
       .TextMatrix(36, 1) = TXTPOLLERALARGO 'Pollera (LARGO)
       .TextMatrix(37, 1) = TXTGRAMAJECUERPO 'CUERPO (Gramaje)
       .TextMatrix(38, 1) = TXTGRAMAJEFONDO 'Fondo (Gramaje)
       .TextMatrix(39, 1) = TXTGRAMAJETAPA 'Tapa (Gramaje)
       .TextMatrix(40, 1) = TXTLAMCUERPOLAMINADO 'CUERPO Laminado
       .TextMatrix(41, 1) = TXTLAMFONDOLAMINADO 'Fondo Laminado
       .TextMatrix(42, 1) = TXT_VPDVDESCARGACM_EDIT 'Cintilla D
       .TextMatrix(43, 1) = TXT_VPDVDESCSEGCM_EDIT 'Cintilla DS
       .TextMatrix(44, 1) = TXT_VPDCargaCm_EDIT 'Cintilla C
       .TextMatrix(45, 1) = TXTPDOCPDOCUMENTOUBIC 'Porta DOCUMENTO(ubicación)
       .TextMatrix(46, 1) = TXTPDOCPORTADOCUMENTOTIPO 'Porta documento Tipo
       .TextMatrix(47, 1) = TXTPORTADOCUMEDIDA 'Porta documento Cantidad
       .TextMatrix(48, 1) = TXTCOMPARTCOMPARTIMENTADO 'Compartimentado
       .TextMatrix(49, 1) = TXT_VPDLARGOMAMAPAROCM 'Largo del Mamparo
       .TextMatrix(50, 1) = TXTDISTBOCAFDOEIDIT 'Distancia de Boca (Comp.)
       .TextMatrix(51, 1) = TXTDISTBOCAFDOEIDIT '"" 'Distancia de Fondo (Comp.)
       .TextMatrix(52, 1) = TXTLINERLINER 'Liner
       .TextMatrix(53, 1) = Text56 'ALTO (Liner)
       .TextMatrix(54, 1) = Text54 'ANCHO (Liner)
       .TextMatrix(55, 1) = TXTEMBALAJEFARDO 'Embalaje U / Fardo
       .TextMatrix(56, 1) = TXTDETAETIQUETA2 'Detalles de la etiqueta
       .TextMatrix(57, 1) = TXTCAPACIDAD 'CAPACIDAD EN KG
       .TextMatrix(58, 1) = "" 'Cantidad de Orifcios
       If TRIM$(CotiMasDato.TXTHILACHA) = "" Then CotiMasDato.TXTHILACHA.TEXT = "NO"
       .TextMatrix(59, 1) = CotiMasDato.TXTHILACHA.TEXT
       If TRIM$(CotiMasDato.TXTDETEMETAL) = "" Then CotiMasDato.TXTDETEMETAL.TEXT = "Ante evento"
       .TextMatrix(60, 1) = CotiMasDato.TXTDETEMETAL.TEXT
'**********************************************************************************
'observaciones TXTOBSAMPLIA
        .TextMatrix(1, 2) = Text33 'Tipo de costura
       .TextMatrix(2, 2) = TXTCOSTURACONPABILO 'Costura con Pabilo
       .TextMatrix(3, 2) = "" 'Tipo de Pabilo
       .TextMatrix(4, 2) = TXTCOSTURAEXTERNA 'Costura Externa
       .TextMatrix(5, 2) = TXTTIPOCOSTURA 'Tipo de costura (V,Plana, zurcida)
       .TextMatrix(6, 2) = TXTESLINGAANCHO ' Eslinga(ANCHO)  'TXTESLINGAANCHO
       .TextMatrix(7, 2) = TXTESLINGAPERIMETRAL 'Eslinga Perimetral
       .TextMatrix(8, 2) = TXTVALVVDDIAMETROCM 'VD (Diámetro)
       .TextMatrix(9, 2) = TXTVDLARGOCM 'VD (Largo)
       .TextMatrix(10, 2) = "" 'VD Material
       .TextMatrix(11, 2) = TXT_VPDVDSLARGOCM 'VDS (Largo)
       .TextMatrix(12, 2) = "" 'Doble Costura VD
       .TextMatrix(13, 2) = TXTPOLLERALAMINADA 'Pollera (Laminada)
       .TextMatrix(14, 2) = TXT_VPDVCGM2 'VC (Gramaje)
       .TextMatrix(15, 2) = TXT_VPDVDGM2 'VD (Gramaje)
       .TextMatrix(16, 2) = TXT_VPDPOLLERAGM2 'Pollera (Gramaje)
       .TextMatrix(17, 2) = TXTTELALAMINADA 'Tapa Laminada
       .TextMatrix(18, 2) = TXT_VPDTELASVALVLAMIN 'Tela Válvulas Láminada
       .TextMatrix(19, 2) = TXT_VPDvcsegcm_EDIT '"" 'Cintilla CS
       .TextMatrix(20, 2) = TXT_VPDPOLLERA_EDIT 'Cintilla pollera
       .TextMatrix(21, 2) = TXTCINTILLALINER_EDIT 'Cintilla Liner
       .TextMatrix(22, 2) = ""
       .TextMatrix(23, 2) = TXTDOCUMENTOUNIDAD 'Porta documento (Ancho)
       .TextMatrix(24, 2) = "" 'Porta documento (LARGO)
       .TextMatrix(25, 2) = "" 'Cantidad de Orifcios
       .TextMatrix(26, 2) = TXT_VPDDIVISIONLATERALES 'División de los laterales (ideal)
       .TextMatrix(27, 2) = TXT_VPDANCHOUTILMAMPARO 'Ancho util del mamparo(ideal)
       .TextMatrix(28, 2) = TXTLINERTIPO 'Tipo
       .TextMatrix(29, 2) = "" ' TXTMICRONAJELINER 'Micronaje (Liner) 'LO PASE MAS ABAJO
       .TextMatrix(30, 2) = TXTUNPALLET.TEXT 'Embalaje U/Pallet 'ESTA POSICION ALGO TIENE EN EL EXCEL QUE NO FUNCIONA
       .TextMatrix(31, 2) = TXT_VPDTELAMAMPARO ' Detalles Mamparo
       .TextMatrix(32, 2) = TXTFUELLELINERCU 'Fuelle (Liner)
       .TextMatrix(33, 2) = TXTMICRONAJELINER
       .TextMatrix(34, 2) = TXTUNPALLET.TEXT 'Embalaje U/Pallet
       .TextMatrix(35, 2) = TXTOBSAMPLIA
       .TextMatrix(36, 2) = TXTFECHA
       .TextMatrix(37, 2) = TXTCANTIUSOS
       .TextMatrix(38, 2) = TXTPRODUCTO 'PRODUCTO A ENVASAR
       '
       'DATOS AUXILIARES
       jj& = 38
       CONDI$ = " CAMPO4  = '" & CStr(FTEC&) & "' AND LISTA = 'DATOSEXT'"
       SQLX$ = "SELECT CAMPO1,CAMPO3 FROM DATASQL WITH(NOLOCK)"
       SQLX$ = SQLX$ + " WHERE " & CONDI$
       SQLX$ = SQLX$ + " ORDER BY DATASQL_ID"
       Set RS = READSET(SQLX$)
       '
       CONTA% = 0
       jj& = jj& + 1
       Do While Not RS.EOF
          CONTA% = CONTA% + 1
          If CONTA% > 6 Then Exit Do 'VA A LLEGAR HASTA LA FILA 44 EN COLUMNAS U y T MANTENER RESERVADAS
          .TextMatrix(jj&, 2) = SAFETEXT(RS!CAMPO1)
          .TextMatrix(jj&, 3) = SAFETEXT(RS!CAMPO3)
          jj& = jj& + 1
          RS.MoveNext
       Loop
       .TextMatrix(jj&, 2) = SAFETEXT(RS!CAMPO1)
       .TextMatrix(jj&, 3) = SAFETEXT(RS!CAMPO3)

       RS.Close
       Set RS = Nothing
       '
    End With
    'OJO SI SE LLAMA A PANTALLA A VENTABNEW CUANDO SE ACEPTA BORRA LAS FILAS QUE ESTAN VACIAS
'    VENTABNEW.Show 1
        
    Call Param_Print_Excel(RUTAEXCEL$, "TODOS LOS DATOS", VENTABNEW.MSF, PUERTOIMPRE$, "1;2;3", 1, 19, "", 1, 0)                     'imprime en excel la 1,2 Y 3 columna de mi msf en la columna 1 a partir de la fila 11 en el excel (este graba E imprime)
    Unload VENTABNEW

    Call SETPRINPORT("RESTORE")





End Sub

Private Sub mnuMarcaNoMostrar_Click(Index As Integer)
   If DERACCE%("Marcar F.Técnicas o Cotizaciones a No Mostrar", "MARNOMOS") < 2 Then Exit Sub

    If Index = 0 Then
10     NRO = SELEC_RECETA(0, 0, , 0, 0, 1)
       If NRO > 0 Then
            CONDI$ = "NUMRECETACTP = " & NRO
            Call ACTUREGISTRO("RECETACTP", "NoMostEnList", CONDI$, "1", REGAF&, "NOMESS")
            FRMZELECHO.TXTBUSCAR = ""
            GoTo 10
       End If
    ElseIf Index = 1 Then
20     NRO = SELEC_RECETA(0, 0, , 0, 0, -1)
       If NRO > 0 Then
           CONDI$ = "NUMRECETACTP = " & NRO
           Call ACTUREGISTRO("RECETACTP", "NoMostEnList", CONDI$, "0", REGAF&, "NOMESS")
           FRMZELECHO.TXTBUSCAR = ""
           GoTo 20
       End If
    End If
    
    '**** HAY QUE VER COMO FILTRAR POR LA MARCA ESTA POR QUE LA LISTA SE ARMA CON UNA VISTA *****
    Dim CTP As New Conteplast
    If Index = 3 Then

30  CONDI$ = " AND (FT_STAT < 9 OR FT_STAT IS NULL) AND COT_NROCOT > 0 AND  (COT_NoMostEnList < 1 OR COT_NoMostEnList  IS NULL) "
    '
    NROFT = CTP.SELEC_FTECNICAYCOTI(CONDI$)
    NROCOT& = XVALO(RESP_ZELE_VECT(3))

    If NROCOT& > 0 Then
            CONDI$ = "NROCOTIZACION = " & NROCOT&
            Call ACTUREGISTRO("COTIZACTP", "NoMostEnList", CONDI$, "1", REGAF&, "NOMESS")
            
            Call ACTUREGISTRO("PRESUVEN", "MARNOVIG", "NROCOTI  = " & NROCOT&, "*", REGAF&, "NOMESS")
            FRMZELECHO.TXTBUSCAR = ""
            GoTo 30
       End If
    ElseIf Index = 4 Then
40  CONDI$ = " AND (FT_STAT < 9 OR FT_STAT IS NULL) AND COT_NROCOT > 0 AND  COT_NoMostEnList =1 "
    '
    NROFT = CTP.SELEC_FTECNICAYCOTI(CONDI$)
    NROCOT& = XVALO(RESP_ZELE_VECT(3))
       If NROCOT& > 0 Then
           CONDI$ = "NROCOTIZACION = " & NROCOT&
           Call ACTUREGISTRO("COTIZACTP", "NoMostEnList", CONDI$, "0", REGAF&, "NOMESS")
           Call ACTUREGISTRO("PRESUVEN", "MARNOVIG", "NROCOTI  = " & NROCOT&, "", REGAF&, "NOMESS")
           FRMZELECHO.TXTBUSCAR = ""
           GoTo 40
       End If
    End If
    Set stp = Nothing

    
End Sub

Private Sub mnuModificarFTG_Click()
       FIRMAELEC.Text1 = ""
       FIRMAELEC.Text2 = ""
       FIRMAELEC.Label1 = ""
    
       FIRMAELEC.Show 1
    
       If FIRMAELEC.ESTADO < 1 Then
          Exit Sub
       Else
          If NIVCONFI%(XVALO(FIRMAELEC.Label1)) < 3 Then
             Call COMUNI("Usuario Con Nivel de Clasificación Inválido Para Esta Operación")
             Exit Sub
          End If
       End If

      Call BLOQUEAR_CONTROLES_ENCONTENEDOR(Me, 1, 1, 1, 1, 0, 1, 0, Me.SSTab1)
      Call BLOQUEAR_CONTROLES_ENCONTENEDOR(Me, 1, 1, 1, 1, 0, 1, 0, Me.Frame1)
      Call BLOQUEAR_CONTROLES_ENCONTENEDOR(Me, 1, 1, 1, 1, 0, 1, 0, Me.Frame2)
      Call VERIFICARSEGUNBIGBAG(Me.TXTCODCONFECCION)
      
      'BLOQUEO LOS TEXT QUE SON TIPO LABEL
      Call BLOQUEAR_CONTROLES_ENCONTENEDOR_Y_COLOR(Me, 2, 1, Me.SSTab1, &H80000002)
      Call BLOQUEAR_CONTROLES_ENCONTENEDOR_Y_COLOR(Me, 2, 1, Me.SSTab1, &HC0C0C0)
      Call BLOQUEAR_CONTROLES_ENCONTENEDOR_Y_COLOR(Me, 2, 1, Me.SSTab1, &H80FF80)
      Call BLOQUEAR_CONTROLES_ENCONTENEDOR_Y_COLOR(Me, 2, 1, Me.SSTab1, &H808080)

      mnuGuardar.Enabled = True
      Command4.Enabled = True
      Command39.Enabled = False

End Sub

Private Sub mnunewCotizacion_Click()
    'DESHABILITE ESTA OPCION PARA QUE SE HAGA DESDE EL BOTON DE SOLICITAR COTIZACIÓN Y LA MISMA SE GUARDE
    'CON LA CANTIDAD .
     Call COMUNI("Opción Deshabilitada, Debe Solicitar Cotización y Luego Seleccionar de la Lista de Pendienes")
     Exit Sub
     '**************************************************************************************
     
     MSF1.Rows = 1
     COD$ = ListaSeleccionArticulos("", ActualizaListaSeleccion, "TIMATE = " & 5)
     ActualizaListaSeleccion = False
     If COD$ <> "" Then
       Call BLANFORMU
       TXTCODIGO = COD$
     Else
       Exit Sub
     End If

     'AGREGA EL NRO CORRESPONDIENTE DE FICHA TECNICA DEL CODIGO.
     Text5 = ""
     Text5 = FICHATECNCIA_SEGUN_CODIGO(CODINT(COD$))
     LBLNROSOLICITUDCOTIZACION = ""
     lblNuevaCotizacion.Visible = True
     TXTNUMCOTIZACION = NumNuevaCotiza
      
'     Call CARGARCOTIZACIONDEFICHAACTIVA
     Call CARGAMSF1(1)
     Call CARGAMSF2(1)
     Call CARGAMSF3(1)
     Call CARGAMSF4(1)
     Call CARGAMSF5(1)
     Call FRMCOTI2.CARGATEXTBLOQUEETIQUETAS(1)
     Call FRMCOTI2.CALCULARCOSTOMATERIALES
     Call FRMCOTI2.CALCULARCOSTODIRECTOBOLSON
     Call CALCULARPORECENTAJES
     FRMCOTI2.Text15 = FECHATEX$(HOY(HO$))
     SSTab1.Tab = 2
End Sub
Sub CALCULARPORECENTAJES()

End Sub


Private Sub mnunewversxcliente_Click()
   NCLIE = ListaSeleccionClientes("", True)
   If NCLIE < 1 Then Exit Sub
   '
   NC = NCLIE
   Text1.TEXT = TRIM$(Str$(NC))

   NRO = SELEC_RECETA(0, 0, NC)
   If NRO > 0 Then
     OPX% = COMALTER%("Confirmación de Grabación de Nueva Ficha Técnica en Base a la Ficha Técnica del Código Seleccionado\\Cancelar\Si Grabar")
     If OPX% > 1 Then
         Text5 = ""
         Text5 = NRO
         LBLNUEVAFICHATECNICA.Visible = True
         TXTFECHA = FECHATEX$(HOY(HO$))
         Me.TXTNUEVAVERSIONDE = Me.TXTCODIGO
         Me.TXTCODIGO = ""
         TXTDESCRICOD = ""
         TXTFECHPREPARO = ""
         TXTAPROBO = ""
         TXTFECHAPROBO = ""

         Call MnuGuardar_Click
         LBLNUEVAFICHATECNICA.Visible = False
         MODIFIC% = 0
         Call VALORESPORDEFECTO
         Call BLOQUEAR_CONTROLES_ENCONTENEDOR(Me, 2, 1, 1, 1, 0, 1, 0, Me.Frame1)
         NUS% = USNBR% Mod 100
         Text4 = LOGVEND.VendeVsUsuario(NUS%)
         
     End If
   End If
   MODIFIC% = 0

End Sub

Private Sub mnuNivelesAprobacion_Click()
    
   USER1% = USNBR% Mod 100
   If USER1% <> 1 Then
      Call COMUNI("Solo Puede Ingresar el Supervisor del Sistema")
      Exit Sub
   End If
   
   Call Mostrar_Datusers
End Sub

Private Sub mnuPedidos_Click()
   Call CONECRUN("PEDIDO07", "Pedidos de Cliente")
End Sub

Private Sub mnuPrintCotiAct_Click()

End Sub

Private Sub mnusalir_Click()
  Unload Me
End Sub

Private Sub mnuSeleccionCarpetaImagen_Click()
    If DERACCE%("CarpImag", "Asignación de Carpeta de Imagenes") < 2 Then Exit Sub
    SELCARPETA.CONTROL1 = "CARPIMAG"
    SELCARPETA.CONTROL2 = "UBICIMAG"
    SELCARPETA.Show 1
    

End Sub

Private Sub mnuseleccodigo_Click()
   NRO = SELEC_RECETA(0, 0, , 1)
   If NRO > 0 Then
     OPX% = COMALTER%("Confirmación de Grabación de Nueva Ficha Técnica en Base a la Ficha Técnica del Código Seleccionado\\Cancelar\Si Grabar")
     If OPX% > 1 Then
         NUS% = USNBR% Mod 100
         Text4 = LOGVEND.VendeVsUsuario(NUS%)
         Text5 = ""
         Text5 = NRO
         LBLNUEVAFICHATECNICA.Visible = True
         TXTFECHA = FECHATEX$(HOY(HO$))
         Me.TXTNUEVAVERSIONDE = Me.TXTCODIGO
         Me.TXTCODIGO = ""
         TXTDESCRICOD = ""
         TXTFECHPREPARO = ""
         TXTAPROBO = ""
         TXTFECHAPROBO = ""
         Call MnuGuardar_Click
         LBLNUEVAFICHATECNICA.Visible = False
         MODIFIC% = 0
         Call VALORESPORDEFECTO
         Call BLOQUEAR_CONTROLES_ENCONTENEDOR(Me, 2, 1, 1, 1, 0, 1, 0, Me.Frame1)

     
'         Call blanformu
'         LBLNUEVAFICHATECNICA.Visible = True
'         TXTFECHA = FECHATEX$(HOY(HO$))
'         Me.TXTNUEVAVERSIONDE = Me.TXTCODIGO
'         Me.TXTCODIGO = ""
'         TXTDESCRICOD = ""
'         Call mnuGuardar_Click
'         LBLNUEVAFICHATECNICA.Visible = False
'         MODIFIC% = 0
'         Call VALORESPORDEFECTO
'         Call BLOQUEAR_CONTROLES_ENCONTENEDOR(Me, 2, 1, 1, 1, 0, 1, 0, Me.Frame1)
'         MODIFIC% = 0

     End If
   End If
   MODIFIC% = 0



End Sub

Private Sub mnuxSolicituddeCotizacion_Click()
    NRO = NROSOLICITUD(1, 0)
    If NRO < 1 Then Exit Sub
    CANTIX = XVALO(RESP_ZELE_VECT(6))
    CODI$ = RESP_ZELE_VECT(3)
    CI1% = CODINT%(CODI$)
    FICHATEC& = XVALO(RESP_ZELE_VECT(9))
    Text5 = ""
    Text5 = FICHATEC&
   
    'AGREGA EL NRO CORRESPONDIENTE DE FICHA TECNICA DEL CODIGO.
    lblNuevaCotizacion.Visible = True
    LBLNROSOLICITUDCOTIZACION = NRO
    TXTNUMCOTIZACION = NumNuevaCotiza
    
    
    If CI1% > 0 Then
        If CANCOMPO%(CI1%) > 0 Then
           'SI YA TIENE ESTRUCTURA (APROBADA) EL CODIGO A COTIZAR CARGA LA ULTIMA QUE TIENE DE ESE CODIGO
           Call CARGAR_GRILLAS_COTIZACION_YA_TIENE_FORMULA(CI1%)
           Command51.Enabled = False
        Else
            'SI NO TIENE ESTRUCTURA ES POR QUE NO ESTA APROBADA.
            Call CARGAMSF1(1)
            Call CARGAMSF2(1)
            Call CARGAMSF3(1)
            Call CARGAMSF4(1)
            Call CARGAMSF5(1)
            Command51.Enabled = True
        End If
        
    End If
    Call FRMCOTI2.TXTCOSTOMATERIAL_Change

    Call FRMCOTI2.CARGATEXTBLOQUEETIQUETAS(1)
    Call FRMCOTI2.CALCULARCOSTOMATERIALES
    Call FRMCOTI2.CALCULARCOSTODIRECTOBOLSON
    FRMCOTI2.TXTCANTSOLIC = ROUND(CANTIX, 0)
    
    Call CALCULARPORECENTAJES
    FRMCOTI2.Text15 = FECHATEX$(HOY(HO$))
    SSTab1.Tab = 2

End Sub

Private Sub MSF1_DblClick()
    REG& = MSF1.MouseRow
    If REG& < 1 Then Exit Sub
    CI1% = CODINT(TXTCODIGO)
    If CI1% > 0 Then
       If FRMCOTI2.STATUSCOTIZA(XVALO(TXTNUMCOTIZACION)) > 0 Then
          Call COMUNI("Atención !!!! Imposible Modificar Formula yá Aprobada\Si Debe Realizar Cambios Recomendamos Generar Nueva Versión")
          Exit Sub
       End If
       If CANCOMPO%(CI1%) > 0 Then
            OPX% = COMALTER%("Formula yá Aprobada\Si Debe Realizar Cambios Recomendamos Generar Nueva Versión\\Cancelar\Corregir (Solo C/Derechos)")
            If OPX% < 2 Then
              Exit Sub
            End If
            If DERACCE%("MOCOTAPR", "Modificar Cotización Aprobada") < 2 Then Exit Sub
       End If
    End If
    
    Unload VENTABNEW
    Campos$ = "COMPONENTE/A28;" '1A
    Campos$ = Campos$ + " FICHA TECNICA(g-m2)/A24;" '2B
    Campos$ = Campos$ + " CODIGO/A16;" '3C
    Campos$ = Campos$ + " DESCRIPCION/A24;" '4D
    Campos$ = Campos$ + " MODULO/F12;" '5E
    Campos$ = Campos$ + " CORTE/F10.2;" '6F
    Campos$ = Campos$ + " CONSUMO (cm)/f12.2;" '7G
    Campos$ = Campos$ + " PRECIO (u$s*M)/F12.4;" '8H
    Campos$ = Campos$ + " PRECIO (u$s)/F12.4;" '9I
    Campos$ = Campos$ + " CORTE X FORMULA/A16" '10J
    Campos$ = Campos$ + " NADA/A0" '11K
    VENTANA.Inicializar = 0
    VENTANA.Campox = Campos$
    'VENTANA.AgregaRegistro = 1
    VENTANA.UtilizaFormula = 1
    VENTANA.HabilitaBorrar = 1
    VENTANA.HabilitaInsertar = 0
    VENTANA.NoMostrarBotonTilde = 0
    VENTANA.TITULO = "GRAMAJE"
    VENTANA.ColEcoCampo(3) = "CODIGO;DESCRIPCION;MASTER;TIMATE <> 5"
'    VENTANA.ColEcoCampo(4) = "CODIGO;DESCRIPCION;MASTER"
'    VENTANA.ColEcoCampo(7) = "NUME_CLI;Fant_Cli;Proved12"
    VENTANA.CampEditables = "3;6"
    
    VENTANA.Formula(7) = "5;*;6"
    VENTANA.Formula(8) = "=COSTOREPOSICION#;3;2"
    'VENTANA.formula(9) = "7;*;8"
    VENTANA.Formula(9) = "=EXCEL;(H1*G1)/100" 'CONSUMO EN CM POR PRECIO X M Y LO DIVIDE POR 100
'    VENTANA.Condicion2(4) = "FAMILIA = COL@2"
    VENTANA.Inicializar = 1
    
    Call CopiarMsfAOtroMsf(Me.MSF1, VENTABNEW.MSF, 1)
    '****************************************************************
    Call T_Espera(1 / 10) ' PARA QUE EL DOBLE CLICK NO APLIQUE SOBRE LA VENTANA QUE SE CARGA
    
50  VENTABNEW.Show 1
    If VENTABNEW.APROBADO% > 0 Then
       RESULTADODIFIERECORTE& = DIFIERECORTE&(VENTABNEW.MSF, 6, 10)
       If RESULTADODIFIERECORTE& > 0 Then
          Call COMUNI("Atención Se Detecta Diferencia en Fila: " & RESULTADODIFIERECORTE& & "\El Valor Original x Formula es de : " & FORMATNUM$(XVALO(VENTABNEW.MSF.TextMatrix(RESULTADODIFIERECORTE&, 10)), "F10.2") & "\Valor Ingresado: " & FORMATNUM$(XVALO(VENTABNEW.MSF.TextMatrix(RESULTADODIFIERECORTE&, 6)), "F10.2"))
'          GoTo 50 PIDERON QUE SOLO AVISE.
       End If
       For i& = 1 To VENTABNEW.MSF.Rows - 1
          If XVALO(VENTABNEW.MSF.TextMatrix(i&, 5)) > 0 And XVALO(VENTABNEW.MSF.TextMatrix(i&, 9)) < 0.00001 Then
             OPX% = COMALTER%("Atención Falta Importe en fila: " & i& & "\\Corregir\Continuar")
             If OPX% < 2 Then
               GoTo 50
             Else
               Exit For
             End If
          End If
       Next i&
       Call CopiarMsfAOtroMsf(VENTABNEW.MSF, MSF1, 1)
       Call FRMCOTI2.CALCULARCOSTOMATERIALES
       MODIFICCOT1% = 1
    End If
    Set VENTANA = Nothing
    Unload VENTABNEW

End Sub

Private Sub MSF2_DblClick()
    REG& = msf2.MouseRow
    If REG& < 1 Then Exit Sub

    CI1% = CODINT(TXTCODIGO)
    If CI1% > 0 Then
       If FRMCOTI2.STATUSCOTIZA(XVALO(TXTNUMCOTIZACION)) > 0 Then
          Call COMUNI("Atención !!!! Imposible Modificar Formula yá Aprobada\Si Debe Realizar Cambios Recomendamos Generar Nueva Versión")
          Exit Sub
       End If
       If CANCOMPO%(CI1%) > 0 Then
            OPX% = COMALTER%("Formula yá Aprobada\Si Debe Realizar Cambios Recomendamos Generar Nueva Versión\\Cancelar\Corregir (Solo C/Derechos)")
            If OPX% < 2 Then
              Exit Sub
            End If
            If DERACCE%("MOCOTAPR", "Modificar Cotización Aprobada") < 2 Then Exit Sub
           End If
    End If
   
''        If CANCOMPO%(CI1%) > 0 Then
''           Call COMUNI("Atención !!!! Imposible Modificar Formula yá Aprobada\Si Debe Realizar Cambios Recomendamos Generar Nueva Versión")
''           Exit Sub
''        End If
'        If CANCOMPO%(CI1%) > 0 Then
'          If lblNuevaCotizacion.Visible = True Then
'            OPX% = COMALTER%("Formula yá Aprobada\Si Debe Realizar Cambios Recomendamos Generar Nueva Versión\\Cancelar\Corregir (Solo C/Derechos)")
'            If OPX% < 2 Then
'              Exit Sub
'            End If
'            If DERACCE%("MOCOTAPR", "Modificar Cotización Aprobada") < 2 Then Exit Sub
'          Else
'            Call COMUNI("Atención !!!! Imposible Modificar Formula yá Aprobada\Si Debe Realizar Cambios Recomendamos Generar Nueva Versión")
'            Exit Sub
'          End If
'        End If
'    End If

    Unload VENTABNEW
    Campos$ = "COMPONENTE/A28;"
    Campos$ = Campos$ + " F.T./A16;"
    Campos$ = Campos$ + " CODIGO/A16;"
    Campos$ = Campos$ + " DESCRIPCION/A24;"
    Campos$ = Campos$ + " MODULO/F12;"
    Campos$ = Campos$ + " CORTE/F10.2;"
    Campos$ = Campos$ + " CONSUMO (cm)/f12.2;"
    Campos$ = Campos$ + " PRECIO (u$s*m)/F12.4;"
    Campos$ = Campos$ + " PRECIO (u$s)/F12.4;"
    Campos$ = Campos$ + " ID/A0"
    VENTANA.Inicializar = 0
    VENTANA.Campox = Campos$
    'VENTANA.AgregaRegistro = 1
    VENTANA.UtilizaFormula = 1
    VENTANA.HabilitaBorrar = 1
    VENTANA.HabilitaInsertar = 0
    VENTANA.NoMostrarBotonTilde = 0
    VENTANA.TITULO = "CINTILLA"
    VENTANA.ColEcoCampo(3) = "CODIGO;DESCRIPCION;MASTER;TIMATE <> 5"

    VENTANA.CampEditables = "3;6"
'    VENTANA.formula(15) = "=EXCEL;(M1/L1)*(M1/L1)*N1"
    VENTANA.Formula(7) = "5;*;6"
'    VENTANA.formula(9) = "7;*;8"
    VENTANA.Formula(8) = "=COSTOREPOSICION#;3;2" '3 COSTODE REPOSICION SEGUN POSICION 3 QUE ES EL CODIGO EXTERNO Y 2 EL TIPO DE MONEDA
    VENTANA.Formula(9) = "=EXCEL;(G1*H1)/100" 'CONSUMO EN CM POR PRECIO X M Y LO DIVIDE POR 100
'    VENTANA.Condicion2(4) = "FAMILIA = COL@2"
    VENTANA.Inicializar = 1
    
    Call CopiarMsfAOtroMsf(Me.msf2, VENTABNEW.MSF, 1)
    '****************************************************************
    Call T_Espera(1 / 10) ' PARA QUE EL DOBLE CLICK NO APLIQUE SOBRE LA VENTANA QUE SE CARGA
    
    VENTABNEW.Show 1
    If VENTABNEW.APROBADO% > 0 Then
'       MODIFICADO(Index) = 1
       Call CopiarMsfAOtroMsf(VENTABNEW.MSF, msf2, 1)
       Call FRMCOTI2.CALCULARCOSTOMATERIALES
       MODIFICCOT1% = 1
    End If
       
    Set VENTANA = Nothing
    Unload VENTABNEW

End Sub




Private Sub MSF3_DblClick()
    REG& = MSF3.MouseRow
    If REG& < 1 Then Exit Sub

    CI1% = CODINT(TXTCODIGO)
    If CI1% > 0 Then
       If FRMCOTI2.STATUSCOTIZA(XVALO(TXTNUMCOTIZACION)) > 0 Then
          Call COMUNI("Atención !!!! Imposible Modificar Formula yá Aprobada\Si Debe Realizar Cambios Recomendamos Generar Nueva Versión")
          Exit Sub
       End If
       If CANCOMPO%(CI1%) > 0 Then
            OPX% = COMALTER%("Formula yá Aprobada\Si Debe Realizar Cambios Recomendamos Generar Nueva Versión\\Cancelar\Corregir (Solo C/Derechos)")
            If OPX% < 2 Then
              Exit Sub
            End If
            If DERACCE%("MOCOTAPR", "Modificar Cotización Aprobada") < 2 Then Exit Sub
           End If
    End If



    Unload VENTABNEW
    Campos$ = "COMPONENTE/A28;F.T./A16;CODIGO/A16;DESCRIPCION/A24;MODULO/F12;CORTE/F10.2;CONSUMO (cm)/f12.2;PRECIO (u$s*m)/F12.4;PRECIO (u$s)/F12.4;NADA/A0"
    VENTANA.Inicializar = 0
    VENTANA.Campox = Campos$
    'VENTANA.AgregaRegistro = 1
    VENTANA.UtilizaFormula = 1
    VENTANA.HabilitaBorrar = 1
    VENTANA.HabilitaInsertar = 0
    VENTANA.NoMostrarBotonTilde = 0
    VENTANA.TITULO = "Mamparo / Eslinga"
    VENTANA.ColEcoCampo(3) = "CODIGO;DESCRIPCION;MASTER;TIMATE <> 5"

    VENTANA.CampEditables = "3;6"
'    VENTANA.formula(15) = "=EXCEL;(M1/L1)*(M1/L1)*N1"
    VENTANA.Formula(7) = "5;*;6"
    'VENTANA.formula(9) = "7;*;8"
    VENTANA.Formula(8) = "=COSTOREPOSICION#;3;2" '3 COSTODE REPOSICION SEGUN POSICION 3 QUE ES EL CODIGO EXTERNO Y 2 LA MONEDA

    VENTANA.Formula(9) = "=EXCEL;(G1*H1)/100" 'CONSUMO EN CM POR PRECIO X M Y LO DIVIDE POR 100
'    VENTANA.Condicion2(4) = "FAMILIA = COL@2"
    VENTANA.Inicializar = 1
    
    Call CopiarMsfAOtroMsf(Me.MSF3, VENTABNEW.MSF, 1)
    '****************************************************************
    Call T_Espera(1 / 10) ' PARA QUE EL DOBLE CLICK NO APLIQUE SOBRE LA VENTANA QUE SE CARGA
    
    VENTABNEW.Show 1
    If VENTABNEW.APROBADO% > 0 Then
'       MODIFICADO(Index) = 1
       Call CopiarMsfAOtroMsf(VENTABNEW.MSF, MSF3, 1)
       Call FRMCOTI2.CALCULARCOSTOMATERIALES
       MODIFICCOT1% = 1
    End If
    Set VENTANA = Nothing
    Unload VENTABNEW

End Sub






Private Sub MSFlexGrid1_Click()

End Sub

Private Sub MSFlexGrid1_DblClick()

End Sub

Private Sub MSFlexGrid1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub


Private Sub MSF4_DblClick()
    REG& = MSF4.MouseRow
    If REG& < 1 Then Exit Sub
    CI1% = CODINT(TXTCODIGO)
    If CI1% > 0 Then
       If FRMCOTI2.STATUSCOTIZA(XVALO(TXTNUMCOTIZACION)) > 0 Then
          Call COMUNI("Atención !!!! Imposible Modificar Formula yá Aprobada\Si Debe Realizar Cambios Recomendamos Generar Nueva Versión")
          Exit Sub
       End If
       If CANCOMPO%(CI1%) > 0 Then
            OPX% = COMALTER%("Formula yá Aprobada\Si Debe Realizar Cambios Recomendamos Generar Nueva Versión\\Cancelar\Corregir (Solo C/Derechos)")
            If OPX% < 2 Then
              Exit Sub
            End If
            If DERACCE%("MOCOTAPR", "Modificar Cotización Aprobada") < 2 Then Exit Sub
           End If
    End If
    
    Unload VENTABNEW
    Campos$ = "COMPONENTE/A28;F.T./A16;CODIGO/A16;DESCRIPCION/A24;MODULO/F12;CORTE/F10.2;CONSUMO (cm)/f12.2;PRECIO (u$s*m)/f12.4;PRECIO (u$s)/F12.4;;NADA/A0"
    VENTANA.Inicializar = 0
    VENTANA.Campox = Campos$
    'VENTANA.AgregaRegistro = 1
    VENTANA.UtilizaFormula = 1
    VENTANA.HabilitaBorrar = 1
    VENTANA.HabilitaInsertar = 0
    VENTANA.NoMostrarBotonTilde = 0
    VENTANA.TITULO = "HILO "
    VENTANA.ColEcoCampo(3) = "CODIGO;DESCRIPCION;MASTER;TIMATE <> 5"

    VENTANA.CampEditables = "3;6"
'    VENTANA.formula(15) = "=EXCEL;(M1/L1)*(M1/L1)*N1"
    VENTANA.Formula(7) = "5;*;6"
    VENTANA.Formula(8) = "=COSTOREPOSICION#;3;2"

    VENTANA.Formula(9) = "=EXCEL;(G1*H1)/100" 'CONSUMO EN CM POR PRECIO X M Y LO DIVIDE POR 100
    
'    VENTANA.Condicion2(4) = "FAMILIA = COL@2"
    VENTANA.Inicializar = 1
    
    Call CopiarMsfAOtroMsf(Me.MSF4, VENTABNEW.MSF, 1)
    '****************************************************************
    Call T_Espera(1 / 10) ' PARA QUE EL DOBLE CLICK NO APLIQUE SOBRE LA VENTANA QUE SE CARGA
    
    VENTABNEW.Show 1
    If VENTABNEW.APROBADO% > 0 Then
'       MODIFICADO(Index) = 1
       Call CopiarMsfAOtroMsf(VENTABNEW.MSF, MSF4, 1)
       Call FRMCOTI2.CALCULARCOSTOMATERIALES
       MODIFICCOT1% = 1
    End If
    Set VENTANA = Nothing
    Unload VENTABNEW

End Sub


Private Sub MSF5_DblClick()
    REG& = MSF5.MouseRow
    If REG& < 1 Then Exit Sub
    
    CI1% = CODINT(TXTCODIGO)
    If CI1% > 0 Then
       If FRMCOTI2.STATUSCOTIZA(XVALO(TXTNUMCOTIZACION)) > 0 Then
          Call COMUNI("Atención !!!! Imposible Modificar Formula yá Aprobada\Si Debe Realizar Cambios Recomendamos Generar Nueva Versión")
          Exit Sub
       End If
       If CANCOMPO%(CI1%) > 0 Then
            OPX% = COMALTER%("Formula yá Aprobada\Si Debe Realizar Cambios Recomendamos Generar Nueva Versión\\Cancelar\Corregir (Solo C/Derechos)")
            If OPX% < 2 Then
              Exit Sub
            End If
            If DERACCE%("MOCOTAPR", "Modificar Cotización Aprobada") < 2 Then Exit Sub
           End If
    End If
    
    Unload VENTABNEW
    Campos$ = "COMPONENTE/A28;F.T./A16;CODIGO/A16;DESCRIPCION/A24;MODULO/F12;CORTE/F10.2;CONSUMO (cm)/f12.2;PRECIO (u$s*m)/F12.4;PRECIO (u$s)/F12.4;NADA/A0"
    VENTANA.Inicializar = 0
    VENTANA.Campox = Campos$
    VENTANA.UtilizaFormula = 1
    VENTANA.HabilitaBorrar = 1
    VENTANA.HabilitaInsertar = 0
    VENTANA.NoMostrarBotonTilde = 0
    VENTANA.TITULO = "ESLINGA PERIMETRAL"
    VENTANA.ColEcoCampo(3) = "CODIGO;DESCRIPCION;MASTER;TIMATE <> 5"

    VENTANA.CampEditables = "3;6"
'    VENTANA.formula(15) = "=EXCEL;(M1/L1)*(M1/L1)*N1"
    VENTANA.Formula(7) = "5;*;6"
    VENTANA.Formula(8) = "=COSTOREPOSICION#;3;2" '3 COSTODE REPOSICION SEGUN POSICION 3 QUE ES EL CODIGO EXTERNO Y 2 LA MONEDA

    'VENTANA.formula(9) = "6;*;8"
    VENTANA.Formula(9) = "=EXCEL;(F1*H1)/100" 'EN ESTE CASO ES CORTE EN CM POR PRECIO X M Y LO DIVIDE POR 100
    
'    VENTANA.Condicion2(4) = "FAMILIA = COL@2"
    VENTANA.Inicializar = 1
    
    Call CopiarMsfAOtroMsf(Me.MSF5, VENTABNEW.MSF, 1)
    '****************************************************************
    Call T_Espera(1 / 10) ' PARA QUE EL DOBLE CLICK NO APLIQUE SOBRE LA VENTANA QUE SE CARGA
    
    VENTABNEW.Show 1
    If VENTABNEW.APROBADO% > 0 Then
       Call CopiarMsfAOtroMsf(VENTABNEW.MSF, MSF5, 1)
       Call FRMCOTI2.CALCULARCOSTOMATERIALES
       MODIFICCOT1% = 1
    End If
    
    Set VENTANA = Nothing
    Unload VENTABNEW


End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
'If SSTab1.Tab = 2 Or SSTab1.Tab = 3 Then
'   Call COMUNI("Cotización Aún No Disponible")
'   SSTab1.Tab = PreviousTab
'   Exit Sub
'End If
   'SOLO ADMINISTRADOR
   NUS% = USNBR% Mod 100
  
   If NUS% > 1 And SSTab1.Tab > 0 And SSTab1.Tab < 4 Then
      If DerechoSupervisor% < 1 Then
        SSTab1.Tab = 0
        Exit Sub
      End If
   End If
   
   If SSTab1.Tab = 2 And XVALO(TXTNUMCOTIZACION) < 1 Then
      Call COMUNI("Debe Seleccionar una Cotización Existente o dar de Alta una Nueva Desde el Menú")
      SSTab1.Tab = PreviousTab
      Exit Sub
   End If

   FRMCOTI2.Text14 = Me.TXTNUMCOTIZACION
   If SSTab1.Tab = 3 And XVALO(FRMCOTI2.Text14) < 1 Then
      Call COMUNI("Debe Seleccionar una Cotización Existente o dar de Alta una Nueva Desde el Menú")
      SSTab1.Tab = PreviousTab
      Exit Sub
   End If
   
  If PreviousTab = 2 And MODIFICCOT1% > 0 And XVALO(TXTNUMCOTIZACION) > 0 Then
      If FRMCOTI2.STATUSCOTIZA(XVALO(TXTNUMCOTIZACION)) < 1 Then
        OPX% = COMALTER%("Se Ha Modificado Datos de la Cotización(1), Desea Grabar\\Cancelar\Si Grabar")
        If OPX% > 1 Then
           Call Command46_Click
           Call FRMCOTI2.Command47_Click
        End If
      End If
      MODIFICCOT1% = 0
        
   End If
   If PreviousTab = 3 And MODIFICCOT2% > 0 And XVALO(TXTNUMCOTIZACION) > 0 Then
     If FRMCOTI2.STATUSCOTIZA(XVALO(Text14)) < 1 Then
      OPX% = COMALTER%("Se Ha Modificado Datos de la Cotización(2), Desea Grabar\\Cancelar\Si Grabar")
      If OPX% > 1 Then
         Call FRMCOTI2.Command47_Click
         Call Command46_Click
      End If
     End If
     MODIFICCOT2% = 0
   End If

   If SSTab1.Tab = 3 Then
     If lblNuevaCotizacion.Visible Then
        Call COMUNI("Atención Debe Guardar Primero la Cotización 1\Grabe Y Vuelva a Intentar")
       SSTab1.Tab = 2
        Exit Sub
     End If
     Call FRMCOTI2.CALCULARCOSTOMATERIALES
     'ABRE EL FORMULARIO EMERGENTE DE COTI2
     FRMCOTI2.Top = Top + (Height - ScaleHeight - 50) + SSTab1.Top + TXTLABEL(13).Top
     FRMCOTI2.Left = Left + SSTab1.Left + Frame10.Left + (Width - ScaleWidth) / 2
     
     FRMCOTI2.Show 1
     SSTab1.Tab = 2 ' SE VUELVE A TAB 2

   End If
   If SSTab1.Tab = 4 Then
     'ABRE EL FORMULARIO EMERGENTE DE COTI2
     PlaniVent.Top = Top + (Height - ScaleHeight - 50) + SSTab1.Top + TXTLABEL(13).Top
     PlaniVent.Left = Left + SSTab1.Left + Frame10.Left + (Width - ScaleWidth) / 2
     PlaniVent.Show 1
     SSTab1.Tab = 0 ' SE VUELVE A TAB 0

   End If
 
End Sub

Private Sub Text1_DblClick()
   If Command1.Enabled = True Then
        Text1 = ""
        TXTRAZOCLI = ""
   End If
   If Control("BASEWALD", "ESTAGRAB") = "SI" Then
      Call COMUNI("Base de Datos en Actualización\Aguarde un Momento y Vuelva a Intentarlo")
      Exit Sub
   End If
   
20 NCLIENTE$ = TRIM$(InputBox$("Ingrese N° de Cliente (Sistema de Envaplast) ", "Alta Cliente", NCX$))
   If NCLIENTE$ = "" Then Exit Sub
   If XVALO(NCLIENTE$) < 1 Then
      Call COMUNI("INGRESO NO VÁLIDO")
      NCX$ = NCLIENTE$
      GoTo 20
   End If
   '
   NC& = XVALO(NCLIENTE$)
   Dim env As New Envap
   Call env.ALTA_CLIENTE_WALDBOT(NC&, 1, RET_OK%)
   If RET_OK% >= 0 Then
     Text1 = NC&
     VendMaestroyWaldbott% = VENDCLI%(NC&)
     NUS% = USNBR% Mod 100
     VendedorActivo% = LOGVEND.VendeVsUsuario(NUS%)
     If VendMaestroyWaldbott% <> VendedorActivo% Then
        OPX% = COMALTER%("Atención !!! El Cliente Dado de Alta Posee Asignado el Vendedor: " & VendMaestroyWaldbott% & "\\Continuar Sin Modificarlo\Asignar Vendedor Activo\Ingresar Manualmente El N° de Vendedor")
        If OPX% = 2 Then 'SI ASIGNA EL VENDEDOR ACTIVO
          VEND% = VendedorActivo%
          Call ACTUREGISTRO("DEUDOR12", "VEND_CLI", "NUME_CLI=" & NC&, VEND%, REGAF&, "NOMESS")
        ElseIf OPX% = 3 Then
          VEND% = VendedorActivo%
50        VENDINGRE$ = TRIM$(InputBox$("Informar", "Ingrese N° de Vendedor a Asignar", VendedorActivo%))
          If XVALO(VENDINGRE$) > 0 Then
             If TRIM$(ECOARCH$("VENDEDOR", Chr$(XVALO(VENDINGRE$)))) = "" Then
                Call COMUNI("Vendedor Inválido")
                GoTo 50
             Else
               VEND% = XVALO(VENDINGRE$) 'SI INGRESA UN VENDEDOR MANUALMENTE
               Call ACTUREGISTRO("DEUDOR12", "VEND_CLI", "NUME_CLI=" & NC&, VEND%, REGAF&, "NOMESS")
             End If
          Else
             Call COMUNI("Se va a Mantener El Vendedor: " & VendMaestroyWaldbott% & " (Para el Cliente Dado de Alta)")
          End If
        End If
        
     End If
   End If
End Sub
 





















Private Sub Text2_Change()
MODIFIC% = 1
End Sub

Private Sub Text29_Change()
MODIFIC% = 1
End Sub

Private Sub Text29_LostFocus()
Call VALORESPORDEFECTO
End Sub

Private Sub Text31_Change()
MODIFIC% = 1
End Sub

Private Sub Text31_LostFocus()
Call VALORESPORDEFECTO
End Sub

Private Sub Text32_Change()
   MODIFIC% = 1
   If TRIM$(Text32) = "" Then
     Text7.Locked = False
     Text7 = ""
     Text7.BackColor = &H80FFFF
   Else
     Text7.Locked = True
     Text7 = ""
     Text7.BackColor = &HC0C0C0
     Call TEXTOLOTES(Text32, ";")
     On Error Resume Next
     If UBound(CADENATEX$) > 0 Then
        Text7 = UBound(CADENATEX$)
     Else
        Text7 = ""
     End If
   End If
End Sub

Private Sub Text32_DblClick()
   Text32 = ""
   Text7 = ""
   Text7.BackColor = &H80FFFF
   Command52.Enabled = True
   
End Sub


Private Sub Text32_LostFocus()
Call VALORESPORDEFECTO
End Sub

Private Sub Text54_Change()
MODIFIC% = 1
End Sub

Private Sub Text56_Change()
MODIFIC% = 1
End Sub


















Private Sub todasAprobadasSinFiltroVend_Click()
   If VALIDASISEMODIFICO% < 0 Then Exit Sub

   NUS% = USNBR% Mod 100
   VENDE% = LOGVEND.VendeVsUsuario(NUS%)

   NRO = SELEC_RECETA(0, 0, , 1, VENDE%, 1)
   If NRO > 0 Then
     Text5 = ""
     LBLNUEVAFICHATECNICA.Visible = False
     Text5 = NRO
     Call VALORESPORDEFECTO
   End If
   MODIFIC% = 0
   'SIEMPRE QUEDA HABILITADA LA POSIBILIDAD DE AGREGAR UN ARCHIVO
   Command47.Enabled = True
   Command43.Enabled = True
   Command45.Enabled = True
   Combo2.Enabled = True

End Sub

Private Sub TXT_VPDANCHOUTILMAMPARO_Change()
MODIFIC% = 1
End Sub

Private Sub TXT_VPDCargaCm_EDIT_Change()
   MODIFIC% = 1
End Sub

Private Sub TXT_VPDCargaCm_EDIT_LostFocus()
Call VALORESPORDEFECTO

End Sub


Private Sub TXT_VPDDISTANCIABOCAFONDOCM_Change()
MODIFIC% = 1
End Sub

Private Sub TXT_VPDDIVISIONLATERALES_Change()
MODIFIC% = 1
End Sub

Private Sub TXTDISTBOCAFDOEIDIT_LostFocus()
    Call VALORESPORDEFECTO
End Sub

Private Sub TXTFUELLELINERCU_Change()
MODIFIC% = 1
End Sub

Private Sub TXT_VPDLARGOMAMAPAROCM_Change()
MODIFIC% = 1
End Sub

Private Sub TXT_VPDPOLLERA_Change()
MODIFIC% = 1
End Sub

Private Sub TXT_VPDPOLLERA_EDIT_Change()
   MODIFIC% = 1
End Sub

Private Sub TXT_VPDPOLLERA_EDIT_LostFocus()
Call VALORESPORDEFECTO

End Sub


Private Sub TXT_VPDPOLLERAGM2_Change()
Call VALORESPORDEFECTO
End Sub

Private Sub TXT_VPDTELASVALVLAMIN_Change()
MODIFIC% = 1
End Sub


Private Sub TXT_VPDVCGM2_Change()
MODIFIC% = 1
End Sub

Private Sub TXT_VPDVCGM2_LostFocus()
    Call VALORESPORDEFECTO

End Sub


Private Sub TXT_VPDvcsegcm_EDIT_Change()
   MODIFIC% = 1
End Sub

Private Sub TXT_VPDvcsegcm_EDIT_LostFocus()
Call VALORESPORDEFECTO

End Sub


Private Sub TXT_VPDVDESCARGACM_EDIT_Change()
MODIFIC% = 1
End Sub

Private Sub TXT_VPDVDESCARGACM_EDIT_LostFocus()
Call VALORESPORDEFECTO

End Sub


Private Sub TXT_VPDVDESCSEGCM_Change()
MODIFIC% = 1
End Sub


Private Sub TXT_VPDVDESCSEGCM_EDIT_Change()
   MODIFIC% = 1
End Sub

Private Sub TXT_VPDVDESCSEGCM_EDIT_LostFocus()
Call VALORESPORDEFECTO

End Sub


Private Sub TXT_VPDVDGM2_Change()
Call VALORESPORDEFECTO
End Sub

Private Sub TXT_VPDVDSLARGOCM_Change()
MODIFIC% = 1
End Sub

Private Sub TXT_VPDVDSLARGOCM_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If
End Sub

Private Sub TXT_VPDVDSLARGOCM_LostFocus()
   Call VALORESPORDEFECTO
End Sub


Private Sub TXTANCHOCM_Change()
MODIFIC% = 1
End Sub

Private Sub TXTAPROBO_DblClick()
   NC = XVALO(Text1)
   If NC < 1 Then
      Call COMUNI("Para Aprobar una Ficha Técnica Debe Tener Asignado un Cliente Destino")
      Exit Sub
   End If
   '
   If TRIM$(Me.TXTCODIGO) <> "" Then
      Call COMUNI("Código Activo Posee Aprobación")
      Exit Sub
   End If
   
   FIRMAELEC.Text1 = ""
   FIRMAELEC.Text2 = ""
   FIRMAELEC.Label1 = ""

   FIRMAELEC.Show 1

   If FIRMAELEC.ESTADO < 1 Then
      Exit Sub
   Else
      If NIVCONFI%(XVALO(FIRMAELEC.Label1)) < 3 Then
         Call COMUNI("Usuario Con Nivel de Clasificación Inválido Para Esta Operación")
         Exit Sub
      End If
      
'      OPX% = COMALTER%("Atención Al \Se Actualizará Toda la Información Tal Cual Esta en Pantalla\\Mantener Información Original\Actualizar")
'      If OPX% < 2 Then Exit Sub
      Screen.MousePointer = 11
      Me.TXTAPROBO = USERNAM$(XVALO(FIRMAELEC.Label1))
      Me.TXTFECHAPROBO = FECHATEX$(HOY(HO$))
      
      If TRIM$(TXTCODIGO) <> "" Then
         Call MnuGuardar_Click
         Screen.MousePointer = 1
         Exit Sub
      End If
      '
      '(largo ancho y alto)
      
      Largo$ = TRIM(Left(TXTCUERPO_LARGOCM, 3))
      ANCHO$ = TRIM(Left(TXTANCHOCM, 3))
      ALTO$ = TRIM(Left(TXTCUERPOALTOCM, 3))

      Largo$ = FORMATOCON0$(XVALO(Largo$), 3)
      ANCHO = FORMATOCON0$(XVALO(ANCHO$), 3)
      ALTO = FORMATOCON0$(XVALO(ALTO$), 3)

      CI1% = 0: DESCRICODI$ = Left(TXTCODCONFECCION.TEXT & " Contenedor Flexible (" & Largo$ & "x" & ANCHO$ & "x" & ALTO$ & ")", 64)
      ESNUEVAVERSION% = 0
      If NC = 99000 Then
         NUEVOCODIGO$ = CREARCODIGO99000$(NC) ' CODIGO CLIENTE 99000
         GoTo 50
      End If
      If TRIM$(TXTNUEVAVERSIONDE) <> "" Then
          CI1% = CODINT%(Me.TXTNUEVAVERSIONDE)
          If CI1% = 0 Then
           Call COMUNI("Código De Referencia para Generar Nueva Versión Inexistente o Inválido")
           Exit Sub
          Else
           ESNUEVAVERSION% = 1
           DESCRICODI$ = DESCRIT0$(CI1%)
          End If
      Else
          'CASO NO ES VERSION DE
          NUEVOCODIGO$ = FORMATOCODIGOCONTEPLAST(NC, 0, 0)
      End If
      '
      If ESNUEVAVERSION% = 1 Then
         NUEVOCODIGO$ = FORMATOCODIGOCONTEPLAST(NC, CI1%, 1)
      End If
      
50    COD$ = TRIM$(NUEVOCODIGO$)
      EXISTECOD% = CODINT%(COD$)
      If EXISTECOD% <> 0 Then
          Call COMUNI("Imposible dar de Alta - Código Existente")
          Exit Sub
      End If
      '
      LONCODI% = XVALO(Control$("MASTER", "LOMAXCOD"))
      If LONCODI% < 15 Or LONCODI% > 24 Then LONCODI% = 15
      VDEF$ = String$(128, 0)
      CI_NEW% = CODINT_NEW
      Call REPLA(VDEF$, MKI$(CI_NEW%), 1, 2)
      Call REPLA(VDEF$, MKI$(GRUCOVEI%(CI_NEW%)), 3, 2)
      Call REPLA(VDEF$, DESCRICODI$, 5, 64)
      Call REPLA(VDEF$, COD$, 77, LONCODI%)
      Call REPLA(VDEF$, "U.", 109, 2)
      Call REPLA(VDEF$, Chr$(1), 101, 1)
      '
      OKX% = 0
'      On Error GoTo ERROR_GUARDAR comentado por que no estaba dando el alta
'      FLEXCONN.BeginTrans

      Call ALTACODIGO(VDEF$, OKX%)
      If OKX% < 1 Then GoTo 80 'se va
      If OKX% = 1 Then
          CODINUE$ = Mid$(VDEF$, 77, LONCODI%)
          TXTCODIGO.TEXT = CODINUE$
          DatosEspecifCliente$ = TRIM$(Left$(TXTCODESPECIFTECCLTE, 64))
          If DatosEspecifCliente$ <> "" Then
              DatosEspecifCliente$ = "Còd.Esp.: " & DatosEspecifCliente$
          End If
          'GRABO POR QUE EN EL OBJETO NO ENTRAN LOS 64 CARACTERES Y LO GRABA TRUNCADO
          SQLX$ = "UPDATE MASTER SET Descripcion = '" & DESCRICODI$ & "',TRAZABLE = 1, Timate =  5, LINDIVIS =  1 , STOMAX = 1, FAMILIA = 'BB',Descridos = '" & DatosEspecifCliente$ & "'"
          SQLX$ = SQLX$ + " WHERE CODINT = " & CI_NEW%
          FLEXCONN.Execute (SQLX$)
       End If
       FTEC& = XVALO(Text5)
       CONDI$ = "NUMRECETACTP =  " & FTEC&
       SQLX$ = "UPDATE RECETACTP SET CODIMASTER = '" & CODINUE$ & "' ,NUCLIEN = " & NC & "  ,NEWVERSIONDE = '" & CODEXT$(CI1%) & "',"
       SQLX$ = SQLX$ + " APROBO = '" & TXTAPROBO & "' , FEAPROBO = '" & CVDAT(HOY(HO$)) & "' , STATUS = 3 ,CODINT = " & CI_NEW%
       SQLX$ = SQLX$ + " WHERE " & CONDI$
       FLEXCONN.Execute (SQLX$)
       
'       Call ACTUREGISTRO("RECETACTP", "CODIMASTER", CONDI$, CODINUE$, REGAF&, "NOMESS")
'       Call ACTUREGISTRO("RECETACTP", "NUCLIEN", CONDI$, NC, REGAF&, "NOMESS")
       
       Text5 = ""
       Text5 = FTEC&
       If ESNUEVAVERSION% < 1 Then
         CUERPO$ = "Alta de Código"
         Call ANOTAHISMASTER(CI_NEW%, CUERPO$)
       ElseIf ESNUEVAVERSION% = 1 Then
         CUERPO$ = "Alta de Código" + vbCrLf
         CUERPO$ = CUERPO$ + "Nueva Versión de: " & CODEXT$(CI1%)
         Call ANOTAHISMASTER(CI_NEW%, CUERPO$)
         TXTNUEVAVERSIONDE = CODEXT$(CI1%)
       End If
       Call CARGAHISTORIAL(CI_NEW%)
   End If
   '

   '
   'ACTUALIZO LA FICHA TECNICA CON LOS DATOS DEL CLIENTE Y
   
'   Call ACTECOMAS

      '
'      On Error GoTo ERROR_GUARDAR
'      FLEXCONN.BeginTrans
'      Call GUARDARFRMENVAP(1) ' parametro 1 opcional para que no presente el mensaje de terminado que esta dentro de esta rutina
'      Call ACTUALIZA_ESTRUCTURA(CI1%, XVALO(Me.TXTCORTE))
'      Call GuardarCodigoCliente(CI1%, XVALO(TXTOBS(31)), TXTOBS(32))
      
'      Call TXTCODIGO_Change
      
'      Call COMUNI("Detalle y Estructura Actualizado")
   Screen.MousePointer = 1

   Unload FIRMAELEC
   Screen.MousePointer = 1
'   Exit Sub
'
'ERROR_GUARDAR:
'    If Err <> 0 Then
'80     A = Err.Description
'       On Error Resume Next
'       FLEXCONN.RollbackTrans
'       On Error GoTo 0
'       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION")
'       Screen.MousePointer = 1
'       Exit Sub
'    Else
'      FLEXCONN.CommitTrans
'
      Call MnuGuardar_Click
      'REFRESCA LA PANTALLA
      NROFT& = XVALO(Text5)
      If NROFT& > 0 Then
          Text5 = ""
          LBLNUEVAFICHATECNICA.Visible = False
          Text5 = NROFT&
          Call VALORESPORDEFECTO
     End If

'
'    End If
80

End Sub

Private Sub TXTCANTICARAS_Change()
MODIFIC% = 1
End Sub

Private Sub TXTCINTILLALINER_EDIT_Change()
MODIFIC% = 1
End Sub

Private Sub TXTCINTILLALINER_DblClick()
  If Command29.Enabled = True Then TXTCINTILLALINER = ""
End Sub


Private Sub TXTCINTILLALINER_EDIT_LostFocus()
Call VALORESPORDEFECTO
End Sub


Private Sub TXTCODIGO_DblClick()
If Command44.Enabled = True Then TXTCODIGO = ""
End Sub

Private Sub TXTCOMPARTCOMPARTIMENTADO_Change()
MODIFIC% = 1
End Sub














Private Sub TXTCOSTURACONPABILO_Change()
MODIFIC% = 1
End Sub

Private Sub TXTCOSTURACONPABILO_DblClick()
   If Command22.Enabled = True Then TXTCOSTURACONPABILO = ""
End Sub


Private Sub TXTCOSTURACONPABILO_LostFocus()
Call VALORESPORDEFECTO
End Sub


Private Sub TXTCOSTURAEXTERNA_Change()
MODIFIC% = 1
End Sub

Private Sub TXTCOSTURAEXTERNA_DblClick()
   If Command21.Enabled = True Then TXTCOSTURAEXTERNA = ""
End Sub


Private Sub TXTCOSTURAEXTERNA_LostFocus()
Call VALORESPORDEFECTO
End Sub


Private Sub TXTCUERPO_LARGOCM_Change()
MODIFIC% = 1
End Sub

Private Sub TXTCUERPOALTOCM_Change()
MODIFIC% = 1
End Sub

Private Sub TXTDESCRICOD_Change()
MODIFIC% = 1
End Sub

Private Sub TXTDETAETIQUETA2_Change()
MODIFIC% = 1
End Sub

Private Sub TXTDETAETIQUETA2_DblClick()
 If Command31.Enabled = True Then TXTDETAETIQUETA2 = ""
End Sub


Private Sub TXTDETALLEETIQUETA_Change()
MODIFIC% = 1
End Sub

Private Sub TXTDISTBOCAFDOEIDIT_Change()
   Call VERIFICARSEGUNBIGBAG(Me.TXTCODCONFECCION)
   MODIFIC% = 1
End Sub

Private Sub TXTDOCUMENTOUNIDAD_Change()
MODIFIC% = 1
End Sub

Private Sub TXTDOCUMENTOUNIDAD_DblClick()
  If Command34.Enabled = True Then TXTDOCUMENTOUNIDAD = ""
End Sub


Private Sub TXTDOCUMENTOUNIDAD_LostFocus()
Call VALORESPORDEFECTO
End Sub


Private Sub TXTEMBALAJEFARDO_Change()
MODIFIC% = 1
End Sub

Private Sub TXTESLINGAANCHO_Change()
MODIFIC% = 1
End Sub

Private Sub TXTESLINGAANCHO_DblClick()
   If Command24.Enabled = True Then TXTESLINGAANCHO = ""
End Sub


Private Sub TXTESLINGAANCHO_LostFocus()
Call VALORESPORDEFECTO
End Sub


Private Sub TXTESLINGAMEDIDA1CM_Change()
MODIFIC% = 1
End Sub

Private Sub TXTESLINGAMED2CMEDITABLE_Change()
MODIFIC% = 1
End Sub


Private Sub TXTESLINGAOJOCM_Change()
MODIFIC% = 1
End Sub

Private Sub TXTESLINGAPERIMETRAL_Change()
MODIFIC% = 1
End Sub

Private Sub TXTESLINGAPERIMETRAL_DblClick()
  If Command23.Enabled = True Then TXTESLINGAPERIMETRAL = ""
End Sub


Private Sub TXTESLINGAPERIMETRAL_LostFocus()
Call VALORESPORDEFECTO
End Sub
Private Sub TXTFECHA_Change()
   MODIFIC% = 1
End Sub
Private Sub TXTGRAMAJECUERPO_Change()
MODIFIC% = 1
End Sub

Private Sub TXTGRAMAJECUERPO_LostFocus()
   Call VALORESPORDEFECTO
End Sub


Private Sub TXTGRAMAJEFONDO_Change()
MODIFIC% = 1
End Sub

Private Sub TXTGRAMAJETAPA_Change()
MODIFIC% = 1
End Sub

Private Sub TXTIMPRESION_Change()
MODIFIC% = 1
End Sub

Private Sub TXTIMPRESION_LostFocus()
   Call VALORESPORDEFECTO
End Sub


Private Sub TXTJUSTA_Change()
    MODIFIC% = 1

End Sub

Private Sub TXTLABEL_DblClick(Index As Integer)
   If Index = 8 Then
     CotiMasDato.Show 1
   End If
End Sub

Private Sub TXTLAMCUERPOLAMINADO_Change()
MODIFIC% = 1
End Sub

Private Sub TXTLAMFONDOLAMINADO_Change()
MODIFIC% = 1
End Sub

Private Sub TXTLINERLINER_LostFocus()
     Call VALORESPORDEFECTO

End Sub

Private Sub TXTLINERTIPO_Change()
MODIFIC% = 1
End Sub

Private Sub TXTMICRONAJELINER_Change()
   MODIFIC% = 1
End Sub

Private Sub TXTMICRONAJELINER_DblClick()
    If Command30.Enabled = True Then TXTMICRONAJELINER = ""
End Sub


Private Sub TXTMICRONAJELINER_LostFocus()
   Call VALORESPORDEFECTO
End Sub

Private Sub TXTNUEVAVERSIONDE_Change()
      If TRIM$(Me.TXTNUEVAVERSIONDE) <> "" Then
         Call BLOQUEAR_CONTROLES_ENCONTENEDOR(Me, 2, 0, 1, 1, 0, 1, 0, Me.Frame1)
      End If
End Sub

Private Sub TXTNUMCOTIZACION_Change()
    Call LIMPIARTEXTSEGUNCOLOR(FRMCOTI2, 1, &HE0E0E0, &HFFFFFF)               'GRIS CLARO y blanco
 
    FRMCOTI2.Text14 = TXTNUMCOTIZACION
    NROCOT& = XVALO(TXTNUMCOTIZACION)
    FTEC& = XVALO(Text5)
    If NROCOT& < 1 Then Exit Sub
    Dim MSF As msFlexGrid
    For i% = 1 To 5
        REG& = 0
        Select Case i%
          Case 1
            Set MSF = MSF1
            MSF1.Rows = 1
          Case 2
            Set MSF = msf2
            msf2.Rows = 2
          Case 3
            Set MSF = MSF3
            MSF3.Rows = 3
          Case 4
            Set MSF = MSF4
            MSF4.Rows = 1
          Case 5
            Set MSF = MSF5
            MSF5.Rows = 1
        End Select
          
        SQLX$ = "SELECT * FROM COTI1 WITH(NOLOCK)"
        SQLX$ = SQLX$ + " WHERE NROCOTIZACION  = " & XVALO(TXTNUMCOTIZACION)
        SQLX$ = SQLX$ + " AND NAMEMSF = '" & MSF.Name & "' "
        SQLX$ = SQLX$ + " ORDER BY ORDEN"
        Set RS = READSET(SQLX$)
        With RS
          Do While Not .EOF
            REG& = REG& + 1
            CI1% = XVALO(!CODINTELEM)
            If MSF.Rows < REG + 1 Then MSF.Rows = REG& + 1
            MSF.TextMatrix(REG&, 1) = SAFETEXT$(!COMPONENTE)
            MSF.TextMatrix(REG&, 2) = SAFETEXT$(!F_TECNICA)
            MSF.TextMatrix(REG&, 3) = SAFETEXT$(!Codigo)
            'enjuague por que en algunos casos no se refrescaba la descripcion
            Descripcix$ = TRIM$(SAFETEXT$(!Descripcion))
            If Descripcix$ = "" Then Descripcix$ = DESCRIT0$(CI1%)
            If UCase$(Descripcix$) = "MATERIAL NO CODIFICADO" Then Descripcix$ = ""
            '
            MSF.TextMatrix(REG&, 4) = Descripcix$
            MSF.TextMatrix(REG&, 5) = XVALO(!moduxbb)
            MSF.TextMatrix(REG&, 6) = FORMATNUM$(XVALO(!corte), "F12.2")
            MSF.TextMatrix(REG&, 7) = FORMATNUM$(XVALO(!moduxbb) * XVALO(!corte), "F12.4")
            MSF.TextMatrix(REG&, 8) = FORMATNUM$(XVALO(!PRECIOXM), "F12.4")
            
            If MSF.Name = "MSF1" Then 'ESCRIBE LOS DATOS DE CONTROL DEL CORTE QUE VIENEN POR FORMULA
              MSF.TextMatrix(REG&, 10) = FORMATNUM$(XVALO(!RESPACORTE), "F12.2")
            End If
            If MSF.Name = "MSF5" Then
                MSF.TextMatrix(REG&, 9) = FORMATNUM$((XVALO(MSF.TextMatrix(REG&, 6)) * XVALO(!PRECIOXM)) / 100, "F12.4")
            Else
              MSF.TextMatrix(REG&, 9) = FORMATNUM$((XVALO(MSF.TextMatrix(REG&, 7)) * XVALO(!PRECIOXM)) / 100, "F12.4")
            End If

            .MoveNext
          Loop
        End With
    Next i%
    RS.Close
    Set RS = Nothing
    
    SQLX$ = "SELECT * FROM COTI2 WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE NROCOTIZACION  = " & XVALO(TXTNUMCOTIZACION)
    Set RS1 = READSET(SQLX$)
    With RS1
     If Not .EOF Then
        FRMCOTI2.Text9(0) = XVALO(!ETIQ_MODULO)
        FRMCOTI2.Text12(0) = FORMATNUM$(XVALO(!ETIQ_COSTO), "F12.4")

        FRMCOTI2.Text7(1) = SAFETEXT$(!SOBRE_DOCTIPO)
        FRMCOTI2.Text8(1) = SAFETEXT$(!SOBRE_DOCMED)
        FRMCOTI2.Text9(1) = XVALO(!SOBRE_MODULO)
        FRMCOTI2.Text12(1) = FORMATNUM$(XVALO(!SOBRE_COSTO), "F12.4")
        
        FRMCOTI2.Text7(3) = FORMATNUM$(XVALO(!LINER_ANCHO), "F12.4")
        FRMCOTI2.Text8(2) = FORMATNUM$(XVALO(!LINER_ALTO), "F12.4")
        FRMCOTI2.Text12(3) = FORMATNUM$(XVALO(!LINER_COSTO), "F12.4")
        FRMCOTI2.Text9(3) = XVALO(!LINER_MODULO)
        FRMCOTI2.Text11(3) = FORMATNUM$(XVALO(!LINER_KG), "F12.4")
        FRMCOTI2.Text12(2) = FORMATNUM$(XVALO(!IMPRE_COSTO), "F12.4")
        FRMCOTI2.Text9(2) = XVALO(!IMPRE_MODULO)
        FRMCOTI2.Text7(4) = FORMATNUM$(XVALO(!ENVO_EMBF), "F12.4")
        FRMCOTI2.Text9(4) = XVALO(!ENVO_MOD)
        FRMCOTI2.Text12(4) = FORMATNUM$(XVALO(!ENVO_COSTO), "F12.4")
        
        FRMCOTI2.TEXT57(0) = SAFETEXT$(!LIBRE1TXT)
        FRMCOTI2.TEXT57(1) = SAFETEXT$(!LIBRE2TXT)
        FRMCOTI2.TEXT57(2) = SAFETEXT$(!LIBRE3TXT)
        FRMCOTI2.Text13(5) = FORMATNUM$(XVALO(!LIBRE1IMPO), "F12.4")
        FRMCOTI2.Text13(6) = FORMATNUM$(XVALO(!LIBRE2IMPO), "F12.4")
        FRMCOTI2.Text13(7) = FORMATNUM$(XVALO(!LIBRE3IMPO), "F12.4")
        FRMCOTI2.TXTFEENTRECOTI = CVDAT(CVINT(!FECHENTRE))
        FRMCOTI2.TXTSCRAP = FORMATNUM$(XVALO(!SCRAPP), "F12.4")
'        FRMCOTI2.TXTUS(1) = FORMATNUM$(XVALO(!LOGYFLETE), "F12.2")
'        FRMCOTI2.TXTUS(2) = FORMATNUM$(XVALO(!IMPUESTO), "F12.2")
        FRMCOTI2.TXTCOMISION = FORMATNUM$(XVALO(!COMISION), "F12.4")
'        FRMCOTI2.TXTUS(4) = FORMATNUM$(XVALO(!ESTRUCTURA), "F12.2")
'        FRMCOTI2.TXTUS(5) = FORMATNUM$(XVALO(!GASTFINAN), "F12.2")
'        FRMCOTI2.TXTUS(6) = FORMATNUM$(XVALO(!OTROS), "F12.2") '
'        FRMCOTI2.TXTUS(7) = FORMATNUM$(XVALO(!OTRO1), "F12.2") '
'        FRMCOTI2.TXTUS(8) = FORMATNUM$(XVALO(!OTRO2), "F12.2") '
                
        FRMCOTI2.TXTUSS2(0) = FORMATNUM$(XVALO(!COSTTOT), "F12.4") '!COSTTOT = XVALO(TXTUSS2(0))
        FRMCOTI2.TXTPORCENT2(1) = FORMATNUM$(XVALO(!PJEUTILIDAD), "F12.4") '!PJEUTILIDAD = XVALO(TXTPORCENT2(1))
        FRMCOTI2.TXTPORCENT2(4) = FORMATNUM$(XVALO(!ComPorReal), "F12.4")
        FRMCOTI2.TXTUSS2(2) = FORMATNUM$(XVALO(!PREVENTA), "F12.4") '!PREVENTA = XVALO(TXTUSS2(2))
        FRMCOTI2.TXTUSS2(3) = FORMATNUM$(XVALO(!PREVENTAREAL), "F12.4") '!OTRO2 = XVALO(TXTUS(8))
        FRMCOTI2.TXTCOSTOMOBRA = FORMATNUM$(XVALO(!COSTMOBRA), "F12.4")
        FRMCOTI2.TXTCOSTOSFIJOS = FORMATNUM$(XVALO(!COSTOSFIJOS), "F12.4")
        FRMCOTI2.Text15 = FECHATEX$(CVINT(!FECHEMISI))
        FRMCOTI2.Text10(3) = XVALO(!MICRONALJELINER)
        FRMCOTI2.Text1 = SAFETEXT$(!OBSERVACIONES)
        FRMCOTI2.TXTCANTSOLIC = XVALO(!CANTICONFIRMADA)
        FRMCOTI2.Text13(2) = FORMATNUM$(XVALO(!VALIMPRESION), "F12.4")
        FRMCOTI2.TXTCODESPVPD = TXTCODIGO 'se guarda en la tabla recetactp
        FRMCOTI2.TXTPORCENT(1) = FORMATNUM$(XVALO(!PORCENTCOMI), "F5.2")
        FRMCOTI2.TXTVALORLINER = FORMATNUM$(XVALO(!IMPOLINER), "F12.4")
        FRMCOTI2.TXTPRECIOFINAL = FORMATNUM$(XVALO(!PREVENTAFINAL), "F12.4")                'CON LINER
        
        FRMCOTI2.TXTFMAPAGO = SAFETEXT$(!MODOPAGO)
        FRMCOTI2.TXTFORMAENTREGA = SAFETEXT$(!FORMAENTREGA)
        FRMCOTI2.TXTMANTENOFERTA = SAFETEXT$(!MANTOFERTA)
        FRMCOTI2.TXTESPECIFCLIENTE_2 = SAFETEXT(!CODESPECIFINT)
        FRMCOTI2.Text3 = SAFETEXT(!PLAZOENTRE)

        End If
    End With
    RS1.Close
    Set RS1 = Nothing

End Sub

Private Sub TXTOBSAMPLIA_Change()
Call VALORESPORDEFECTO
End Sub

Private Sub TXTPALLET_Change()
MODIFIC% = 1
End Sub

Private Sub TXTPALLET_DblClick()
   If Command32.Enabled = True Then TXTPALLET = ""
End Sub


Private Sub TXTPDOCPDOCUMENTOUBIC_Change()
MODIFIC% = 1
End Sub

Private Sub TXTPDOCPORTADOCUMENTOTIPO_Change()
MODIFIC% = 1
End Sub

Private Sub TXTPOLLERALAMINADA_DblClick()
   If Command27.Enabled = True Then TXTPOLLERALAMINADA = ""
End Sub

Private Sub TXTPOLLERALAMINADA_LostFocus()
Call VALORESPORDEFECTO
End Sub

Private Sub TXTPOLLERALARGO_Change()
MODIFIC% = 1
End Sub










Private Sub TXTPORTADOCUMEDIDA_Change()
MODIFIC% = 1
End Sub

Private Sub TXTPORTADOCUMEDIDA_DblClick()
  If Command35.Enabled = True Then TXTPORTADOCUMEDIDA = ""
End Sub


Private Sub TXTPORTADOCUMEDIDA_LostFocus()
Call VALORESPORDEFECTO
End Sub


Private Sub TXTPRODUCTO_Change()
MODIFIC% = 1
End Sub

Private Sub TXTRAZOCLI_Change()
  A = TXTRAZOCLI
End Sub

Private Sub TXTSTOTSOLAPACOTIZACION_Change()
   FRMCOTI2.TXTTRANSPORTE = TXTSTOTSOLAPACOTIZACION
End Sub




Private Sub TXTTELALAMINADA_Change()
MODIFIC% = 1
End Sub

Private Sub TXTTELALAMINADA_DblClick()
  If Command28.Enabled = True Then TXTTELALAMINADA = ""
End Sub


Private Sub TXTTELALAMINADA_LostFocus()
Call VALORESPORDEFECTO
End Sub


Private Sub TXTTELAVENTILADA_Change()
MODIFIC% = 1
End Sub

Private Sub TXTTIPOCOSTURA_Change()
MODIFIC% = 1
End Sub

Private Sub TXTTIPOCOSTURA_DblClick()
  If Command25.Enabled = True Then TXTTIPOCOSTURA = ""
End Sub


Private Sub TXTTIPOCOSTURA_LostFocus()
Call VALORESPORDEFECTO
End Sub


Private Sub TXTUNPALLET_Change()
MODIFIC% = 1
End Sub











Private Sub TXTVALVVCDIAMETROCM_Change()
MODIFIC% = 1
End Sub

Private Sub TXTVALVVCLARGOCM_Change()
MODIFIC% = 1
End Sub

Private Sub TXTVALVVDDIAMETROCM_Change()
MODIFIC% = 1
End Sub

Private Sub TXTVALVVDDIAMETROCM_DblClick()
  If Command26.Enabled = True Then TXTVALVVDDIAMETROCM = ""
End Sub


Private Sub TXTVALVVDDIAMETROCM_LostFocus()
Call VALORESPORDEFECTO
End Sub


Private Sub TXTVCLARGOCM_Change()
MODIFIC% = 1
End Sub

Private Sub TXTVDLARGOCM_Change()
MODIFIC% = 1
End Sub

Private Sub VScroll1_Change()
   Picture5.Top = (-100) * VScroll1.Value
End Sub

Private Sub VScroll1_Scroll()
   Picture5.Top = (-100) * VScroll1.Value
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   If Me.Visible = True Then
      Call CIERRARCH("*.*")
      Call BAJALDISCO
      Call FINAL("")
   End If
End Sub
Sub VALORESPORDEFECTO()
         'Eslinga (Medida1 Cm) = SI(C16=0,"",+C16/2-10) (VALOR POR DEFECTO SEGUNDA COLUMNA)
'        If TRIM$(TXTCODIGO) <> "" Then Exit Sub
        If XVALO(TXTCUERPOALTOCM) = 0 Then
           TXT_VPDFESLINGA = ""
        Else
          TXT_VPDFESLINGA = XVALO(TXTCUERPOALTOCM) / 2 - 10
        End If
        
        'Eslinga (Medida2 Cm) =SI(C17=0,"",SI(D10<1200,20,25))
        If XVALO(TXTESLINGAOJOCM) = 0 Then
            TXTESLINGAMEDIDA2CM = ""
        ElseIf XVALO(TXTCAPACIDAD) < 1200 Then
            TXTESLINGAMEDIDA2CM = "20"
        Else
            TXTESLINGAMEDIDA2CM = "25"
        End If
        '
        'VC (Diámetro en cm) = SI(C20="NO","",36) (VALOR POR DEFECTO SEGUNDA COLUMNA)
        If XVALO(TXTVALVVCDIAMETROCM) <= 0 Then
            TXT_VPDVCDIAMETROCM = ""
        Else
            TXT_VPDVCDIAMETROCM = 36
        End If
        '
        'VC (Largo en cm) =SI(C21="NO","",50)
'        If XVALO(TXTVCLARGOCM) <= 0 Then
'            TXT_VPDVCLARGOCM = ""
'        Else 'DIJO LAUTARO QUE SE DEBE MOSTRAR SIEMPRE LOS 50 COMO SUGERIDO
            TXT_VPDVCLARGOCM = 50
'        End If
        '
        'VC (g/m2) = SI(C20=0,"",SI(C20="NO","NO","60 / 65 + 10/15"))
        If XVALO(TXTVALVVCDIAMETROCM) = 0 Then
           TXT_VPDVCGM2 = ""
        ElseIf TRIM$(UCase(TXTVALVVCDIAMETROCM)) = "NO" Then
            TXT_VPDVCGM2 = "NO"
        Else
           TXT_VPDVCGM2 = "60 / 65 + 10/15"
        End If
        '
        'VD (g/m2) = SI(G20=0,"",SI(G20="NO","NO","60 / 65 + 10/15"))
        If XVALO(TXTVALVVDDIAMETROCM) = 0 Then
           TXT_VPDVDGM2 = ""
        ElseIf TRIM$(UCase(TXTVALVVDDIAMETROCM)) = "NO" Then
            TXT_VPDVDGM2 = "NO"
        Else
           TXT_VPDVDGM2 = "60 / 65 + 10/15"
        End If
        '
        'Pollera (g/m2) =SI(G23="","",SI(C23=0,"",SI(C23="NO","NO",SI(G23="SI","60 / 75 + 10/15","60/65"))))
        If TRIM$(TXTPOLLERALAMINADA) <> "" Then
          If XVALO(TXTPOLLERALARGO) = 0 Then
             TXT_VPDPOLLERAGM2 = ""
           ElseIf TRIM(TXTPOLLERALARGO) <> "NO" Then
              If TRIM$(TXTPOLLERALAMINADA) = "SI" Then
                 TXT_VPDPOLLERAGM2 = "60 / 75 + 10/15"
              Else
                 TXT_VPDPOLLERAGM2 = "60/65"
              End If
           End If
        End If
        '
        'Tela Válvulas Láminada = SI(C20=0,"",SI(Y(C20="NO",G20="NO"),"","SI"))
        If XVALO(TXTVALVVCDIAMETROCM) = 0 And IsNumeric(TXTVALVVCDIAMETROCM) Then
            TXT_VPDTELASVALVLAMIN = ""
        Else
           If TRIM$(UCase(TXTVALVVCDIAMETROCM)) = "NO" And TRIM$(UCase(TXTVALVVDDIAMETROCM)) = "NO" Then
            TXT_VPDTELASVALVLAMIN = ""
           Else
             TXT_VPDTELASVALVLAMIN = "SI"
           End If
        End If
        '
        'Pollera=SI(C23=0,"",SI(C23="NO","NO",80))
        If XVALO(TXTPOLLERALARGO) = 0 And IsNumeric(TXTPOLLERALARGO) Then
           TXT_VPDPOLLERA = ""
        Else
           If TRIM$(UCase$(TXTPOLLERALARGO)) <> "NO" Then
                TXT_VPDPOLLERA = 80
           End If
        End If
        '
        'C16 ES EL ALTO
        'TXT_VPDLARGOMAMAPAROCM=SI(O(D33="NO",D33=""),"",SI(C16<=130,C16-20,SI(C16>=201,C16-40,C16-30)))
        If TRIM$(UCase$(TXTCOMPARTCOMPARTIMENTADO)) = "NO" Or TRIM$(TXTCOMPARTCOMPARTIMENTADO) = "" Then
           TXT_VPDLARGOMAMAPAROCM = ""
        Else
          If IsNumeric(TXTCUERPOALTOCM) Then
              TXT_VPDLARGOMAMAPAROCM = XVALO(TXTCUERPOALTOCM) - XVALO(TXTDISTBOCAFDOEIDIT) * 2 ' MODIFICADO SEGUN W.S RAUL 15-06-20

              If XVALO(TXTCUERPOALTOCM) <= 130 Then
                  
                  TXT_VPDDISTANCIABOCAFONDOCM = "10"
               Else
                  If XVALO(TXTCUERPOALTOCM) >= 201 Then
'                     TXT_VPDLARGOMAMAPAROCM = XVALO(TXTCUERPOALTOCM) - 40 'MODIFICADO SEGUN W.S RAUL 15-06-20
                     TXT_VPDDISTANCIABOCAFONDOCM = "20"
                  Else
'                      TXT_VPDLARGOMAMAPAROCM = XVALO(TXTCUERPOALTOCM) - 30
                      TXT_VPDDISTANCIABOCAFONDOCM = "15"
                  End If
              End If
           End If
        End If
        '
        'V Descarga (cm)=SI(G20=0,"",SI(G20="NO","",80)) TXT_VPDVDESCARGACM
        If IsNumeric(TXTVALVVDDIAMETROCM) And XVALO(TXTVALVVDDIAMETROCM) = 0 Then
              TXT_VPDVDESCARGACM = ""
        Else
           If TRIM(UCase(TXTVALVVDDIAMETROCM)) = "NO" Then
               TXT_VPDVDESCARGACM = ""
           Else
                TXT_VPDVDESCARGACM = 80
           End If
        End If
        '
        'V Desc Seg (cm) =SI(O(G20="NO",G20=""),"",BUSCARV(G20,'Carga Esp'!K2:L8, 2,0))
        If TRIM$(TXTVALVVDDIAMETROCM) = "NO" Or TRIM$(TXTVALVVDDIAMETROCM) = "" Then
             TXT_VPDVDESCSEGCM = ""
             
        Else
             TXT_VPDVDESCSEGCM = V_Desc_Seg_cm_SEGUN_VD_Diámetro_en_cm(XVALO(TXTVALVVDDIAMETROCM))
        End If
        '
        'V Carga (Cm)=SI(C20=0,"",SI(C20="NO","NO",80))
        If IsNumeric(TXTVALVVCDIAMETROCM) And XVALO(TXTVALVVCDIAMETROCM) = 0 Then
           TXT_VPDCargaCm = ""
        Else
           If TRIM$(TXTVALVVCDIAMETROCM) = "NO" Then
               TXT_VPDCargaCm = "NO"
           Else
               TXT_VPDCargaCm = 80
           End If
        End If
        '
        
        'V C. Seg (Cm)=SI(C22=0,"",SI(C23="NO","NO",BUSCARV(C20,'Carga Esp'!K2:L8,2,0)))
        If IsNumeric(TXTVALVVCLARGOCM) And XVALO(TXTVALVVCLARGOCM) = 0 Then
            TXT_VPDvcsegcm = ""
        Else
            If TRIM$(TXTPOLLERALARGO) = "NO" Then
                TXT_VPDvcsegcm = "NO"
             Else
                TXT_VPDvcsegcm = V_Desc_Seg_cm_SEGUN_VD_Diámetro_en_cm(TXTVALVVCDIAMETROCM)
             End If
        End If
        '
        '
        'TEXT19 =SI(C22=0,"",SI(C22="NO","",SI(C20="NO","",+C20/2)))  (VCS (Largo en cm)


        If IsNumeric(TXTVALVVCLARGOCM) And XVALO(TXTVALVVCLARGOCM) = 0 Then
             Text19 = ""
        Else
         If TRIM$(TXTVALVVCLARGOCM) = "NO" Then
             Text19 = ""
          Else
              If TRIM$(TXTVALVVCDIAMETROCM) = "NO" Then
                    Text19 = ""
               Else
                    Text19 = XVALO(TXTVALVVCDIAMETROCM) / 2
               End If
          End If
        End If
        'TXT_VPDVDSLARGOCM  = SI(G20=0,"",SI(G20="NO","NO",+G20/2))
        If IsNumeric(TXTVALVVDDIAMETROCM) And XVALO(TXTVALVVDDIAMETROCM) = 0 Then
                 TXT_VPDVDSLARGOCM = ""
        Else
                 If TRIM$(UCase$(TXTVALVVDDIAMETROCM)) = "NO" Then
                     TXT_VPDVDSLARGOCM = "NO"
                 Else
                     TXT_VPDVDSLARGOCM = XVALO(TXTVALVVDDIAMETROCM) / 2
                 End If
                 
        End If

        ' Compartimentado= SI(O(D33="NO",D33=""),"",+(C16-D34)/2)
        If TRIM$(TXTCOMPARTCOMPARTIMENTADO) = "NO" Or TRIM$(TXTCOMPARTCOMPARTIMENTADO) = "" Then
'            TXTDISTBOCAFDOEIDIT = ""
            TXT_VPDDISTANCIABOCAFONDOCM = ""
        Else
'            TXTDISTBOCAFDOEIDIT = (XVALO(TXTCUERPOALTOCM) - XVALO(TXT_VPDLARGOMAMAPAROCM)) / 2
'            TXT_VPDDISTANCIABOCAFONDOCM = (XVALO(TXTCUERPOALTOCM) - XVALO(TXT_VPDLARGOMAMAPAROCM)) / 2
        End If
        '
        'División de laterales =SI(O(D33="NO",D33=""),"",CONCATENAR('Carga Esp'!Y2," "," ",'Carga Esp'!Z2," "," ",'Carga Esp'!Y2))

        If TRIM$(TXTCOMPARTCOMPARTIMENTADO) = "NO" Or TRIM$(TXTCOMPARTCOMPARTIMENTADO) = "" Then
           TXT_VPDDIVISIONLATERALES = ""
        Else
           Y2$ = SAFETEXT$(Int(XVALO(TXTANCHOCM) / 3))
           Z2$ = XVALO(TXTANCHOCM) - XVALO(Y2$) * 2
           TXT_VPDDIVISIONLATERALES = Y2$ & " " & Z2$ + " " + Y2$
        End If
       
        'Ancho util del mamparo =SI(O(D33="NO",D33=""),"",ENTERO(RAIZ('Carga Esp'!Y2*'Carga Esp'!Y2+'Carga Esp'!Y2*'Carga Esp'!Y2)-1))
'                                =SI(O(D33="NO",D33=""),"",ENTERO(RAIZ('Carga Esp'!Y2*'Carga Esp'!Y2+'Carga Esp'!Y2*'Carga Esp'!Y2)-1))
        If TRIM$(TXTCOMPARTCOMPARTIMENTADO) = "NO" Or TRIM$(TXTCOMPARTCOMPARTIMENTADO) = "" Then
           TXT_VPDANCHOUTILMAMPARO = ""
        Else
           VVAL = Int(XVALO(TXTANCHOCM) / 3) ' VER ACA EL TEMA DE LOS INT
           Y22 = Int(VVAL * VVAL + VVAL * VVAL)
           Y22 = Int(Y22 ^ (1 / 2)) - 1 'RAIZ CUADRADA
           TXT_VPDANCHOUTILMAMPARO = Y22
        End If
        '
        'TXT_VPDFUELLELINERCU = SI(C36 = "SI", C14 / 2, "")
        If UCase$(TRIM$(TXTLINERLINER)) = "SI" Then
            TXT_VPDFUELLELINERCU = XVALO(TXTANCHOCM) / 2
        Else
            TXT_VPDFUELLELINERCU = ""
        End If
        '
        'TXT_VPDVDDIAMENTRO =SI(G20="NO","",36)  (VD (Diámetro en cm))
        If TRIM$(TXTVALVVDDIAMETROCM) = "NO" Then
           TXT_VPDVDDIAMENTRO = ""
        Else
          TXT_VPDVDDIAMENTRO = 36
        End If
        
        'TXTVPDVDLARGO =SI(G21="NO","",50)   (VD (Largo en cm))
        If TRIM$(TXTVDLARGOCM) = "NO" Then
           TXT_VPDVDLARGO = ""
        Else
          TXT_VPDVDLARGO = 50
        End If
        
        '
        '=
        'TXT_VPDTELAMAMPARO = SI(G34="","",BUSCARV(G34+5,MP,2))
        If TRIM$(TXT_VPDANCHOUTILMAMPARO) = "" Then
          TXT_VPDTELAMAMPARO = ""
        Else
            TXT_VPDTELAMAMPARO = ValorTelaMamparo_SEGUN_AnchoMamparo(XVALO(TXT_VPDANCHOUTILMAMPARO + 5))
            TXT_VPDTELAMAMPARO = "TTL " & TXT_VPDTELAMAMPARO & " 60/65 + 10/15"
        End If
        
        
        
'        'Alto (Liner)= SI(O(C36="NO",C36=""),"",SI(C37="FONDO SELLADO",SI(C20="",C14/2+C16+C23+10,C14+C16+C21+10),SI(C20="",C14/2+C16+C23+10+G21+SI(G21=0,0,10),C14+C16+C21+10+G21+SI(G21=0,0,10))))
'        If UCase$(TRIM$(TXTLINERLINER)) = "NO" Or UCase$(TRIM$(TXTLINERLINER)) = "" Then
'                TXT_VPDALTOLINER = ""
'            Else
'                If UCase$(TRIM$(TXTLINERTIPO)) = "FONDO SELLADO" Then
'                    If TRIM$(TXTVALVVCDIAMETROCM) = "" Then
'                         TXT_VPDALTOLINER = XVALO(TXTANCHOCM) / 2 + XVALO(TXTCUERPOALTOCM) + XVALO(TXTPOLLERALARGO) + 10
'                    Else
'                         TXT_VPDALTOLINER = XVALO(TXTANCHOCM) + XVALO(TXTCUERPOALTOCM) + XVALO(TXTVCLARGOCM) + 10
'                    End If
'                Else
'                   If TRIM$(TXTVALVVCDIAMETROCM) = "" Then
'                       TXT_VPDALTOLINER = XVALO(TXTANCHOCM) / 2 + XVALO(TXTCUERPOALTOCM) + XVALO(TXTPOLLERALARGO) + 10 + XVALO(TXTVDLARGOCM) & "cm"
'                       If XVALO(TXTVDLARGOCM) = 0 Then
'                           TXT_VPDALTOLINER = XVALO(TXT_VPDALTOLINER) + 0
'                        Else
'                           TXT_VPDALTOLINER = XVALO(TXT_VPDALTOLINER) + 10 & "cm"
'                       End If
'                    Else
'                         TXT_VPDALTOLINER = XVALO(TXTANCHOCM) + XVALO(TXTCUERPOALTOCM) + XVALO(TXTPOLLERALARGO) + 10 + XVALO(TXTVDLARGOCM) & "cm"
'                         If XVALO(TXTVDLARGOCM) = 0 Then
'                             TXT_VPDALTOLINER = TXT_VPDALTOLINER + 0
'                         Else
'                             TXT_VPDALTOLINER = TXT_VPDALTOLINER + 10
'                          End If
'                     End If
'                End If
'         End If
         'TXT_VPDANCHO = SI(O(C36 = "NO", C36 = ""), "", (C14 * 2 + 5) - D38 * 2) (Fuelle (Liner) c/u)
         If TRIM$(UCase$(TXTLINERLINER)) = "NO" Or TRIM$(UCase$(TXTLINERLINER)) = "" Then
            TXT_VPDANCHO = ""
         Else
            TXT_VPDANCHO = (XVALO(TXTANCHOCM) * 2 + 5) - (XVALO(TXT_VPDFUELLELINERCU) * 2) & " + fuelles"
         End If
         
 

        If UCase$(TRIM$(TXTLINERLINER)) = "NO" Or UCase$(TRIM$(TXTLINERLINER)) = "" Then
                TXT_VPDALTOLINER = ""
        Else
           If UCase$(TRIM$(TXTLINERTIPO)) = "FONDO SELLADO" Then
              If TRIM$(TXTVALVVCDIAMETROCM) = "" Then
                 TXT_VPDALTOLINER = XVALO(TXTANCHOCM) / 2 + XVALO(TXTCUERPOALTOCM) + XVALO(TXTPOLLERALARGO) + 10
              Else
                  TXT_VPDALTOLINER = XVALO(TXTANCHOCM) + XVALO(TXTCUERPOALTOCM) + XVALO(TXTVCLARGOCM) + 10
              End If
           Else
              If TRIM$(TXTVALVVCDIAMETROCM) = "" Then
                  TXT_VPDALTOLINER = XVALO(TXTANCHOCM) / 2 + XVALO(TXTCUERPOALTOCM) + XVALO(TXTPOLLERALARGO) + 10 + XVALO(TXTVDLARGOCM)
                  If XVALO(TXTVDLARGOCM) = 0 And IsNumeric(TXTVDLARGOCM) Then
                     RESULTADO1 = 0
                  Else
                     RESULTADO1 = 10 * 2 ' RESUELVE EL ULITMO CALCULO +SI(G21=0,0,10))))
                  End If
              Else
                  RESULTADO1 = XVALO(TXTANCHOCM) + XVALO(TXTCUERPOALTOCM) + XVALO(TXTVCLARGOCM) + 10 + XVALO(TXTVDLARGOCM)
              End If
              TXT_VPDALTOLINER = RESULTADO1
           End If
           
       End If
       
       'C16 = TXTCUERPOALTOCM (Alto (Cm))
        'C17 = TXTESLINGAOJOCM  (Eslinga (Ojo Cm))
        'D33 = TXTCOMPARTCOMPARTIMENTADO
        'D34 = TXT_VPDLARGOMAMAPAROCM  (Largo del Mamparo (cm))
        'D38 = TXT_VPDFUELLELINERCU (Fuelle (Liner) c/u)
        'C14 = TXTANCHOCM
        'C16 = TXTCUERPOALTOCM
        'C20 = TXTVALVVCDIAMETROCM
        'C21 = TXTVCLARGOCM
        'C22 = TXTVALVVCLARGOCM
        'C23 = TXTPOLLERALARGO
        'C14 = TXTANCHOCM
        'C36 = TXTLINERLINER
        'C37 = TXTLINERTIPO
        'G20 = TXTVALVVDDIAMETROCM (VD (Diámetro en cm))
        'G21 = TXTVDLARGOCM
        'G23 = TXTPOLLERALAMINADA (Pollera (Laminada))
        'G34 = TXT_VPDANCHOUTILMAMPARO
        
        '=========================================================
        'AGREGADO SEGUN MINUTA 06/12/19
        'SI IMPRESION ES NO BLOQUEA Y LIMPIA
        If TRIM$(UCase$(TXTIMPRESION)) = "NO" Then
            Command6.Enabled = False
            Command52.Enabled = False
            TXTCANTICARAS.Locked = True
            Text29.Locked = True
            Text31.Locked = True
            TXTCANTICARAS = ""
            Text29 = ""
            Text31 = ""
        Else
            Command6.Enabled = True
            Command52.Enabled = True
            TXTCANTICARAS.Locked = False
            Text29.Locked = False
            Text31.Locked = False
        End If
        '
        Call VERIFICARSEGUNBIGBAG(TXTCODCONFECCION)
        
        'AGREGADOS SEGUN MINUTA 03/12/19 (X:\CONTEPLAST\FLEXOFT\PROYECTO)
        '4.-Si Tipo de Costura = 'V' Eslinga perimetral  Bloqueado y Texto en casillero No.
        If TRIM$(TXTTIPOCOSTURA) = "V" Then
           TXTESLINGAPERIMETRAL = "NO"
           TXTESLINGAPERIMETRAL.Locked = True
           Command23.Enabled = False
        Else
           If TRIM$(UCase$(TXTESLINGAPERIMETRAL)) = "NO" Then TXTESLINGAPERIMETRAL = ""
           TXTESLINGAPERIMETRAL.Locked = False
           Command23.Enabled = True
        End If
'        5.-Si en Gramaje.Cuerpo (g/m2) se elige de la lista de selección algun texto que posea
'        el operador '+'  el Lam.Cuerpo Laminado = 'Si' caso contrario NO
         If InStr(TXTGRAMAJECUERPO, "+") > 0 Then
            TXTLAMCUERPOLAMINADO = "SI"
            TXTLAMCUERPOLAMINADO.Locked = True 'POR AHORA LO DEJO SIEMPRE NO EDITABLE HASTA  VER SI VA BIEN
            Command13.Enabled = False
         Else
            TXTLAMCUERPOLAMINADO = "NO"
            Command13.Enabled = False '
            TXTLAMCUERPOLAMINADO.Locked = False
         End If
         
        ' 6.-Si en Gramaje.Fondo(g/m2) se elige de la lista de selección algun texto que posea el operador '+'
        ' el Lam.Fondo Laminado = 'Si' caso contrario NO y lo mismo para Lam.Tapa Laminado.End Sub
         If InStr(TXTGRAMAJEFONDO, "+") > 0 Then
            TXTLAMFONDOLAMINADO = "SI"
            TXTLAMFONDOLAMINADO.Locked = True 'POR AHORA LO DEJO SIEMPRE NO EDITABLE HASTA  VER SI VA BIEN
            Command12.Enabled = False
         Else
            TXTLAMFONDOLAMINADO = "NO"
            Command13.Enabled = True '
            TXTLAMFONDOLAMINADO.Locked = False
         End If

         If InStr(TXTGRAMAJETAPA, "+") > 0 Then
            TXTTELALAMINADA = "SI"
            TXTTELALAMINADA.Locked = True 'POR AHORA LO DEJO SIEMPRE NO EDITABLE HASTA  VER SI VA BIEN
            Command28.Enabled = False
         Else
            TXTTELALAMINADA = "NO"
            Command28.Enabled = True '
            TXTTELALAMINADA.Locked = False
         End If

        '7.-El casillero Compart.Dist. Boca y Fdo.(cm)  se correrá hacia la derecha (valor por formula sugerido)
        ' y se agregara un casillero en lugar de este que deberá ser editable para el ingreso a registrar.
        'REALIZADO
        
        '9.-Si Liner = NO el Tipo y Alto (liner), Micronanje (Liner) y ancho del Grupo Liner deben quedar vacíos
        ' y deshabilitado.
        If TRIM$(UCase$(TXTLINERLINER)) = "NO" Then
            TXTLINERTIPO = ""
            Text56 = ""
            TXTMICRONAJELINER = ""
            Text54 = ""
            Text56.Locked = True
            Text54.Locked = True
            
            TXT_VPDFUELLELINERCU = ""
            Command18.Enabled = False
            Command30.Enabled = False
         Else
            Text56.Locked = False
            Text54.Locked = False
            Command18.Enabled = True
            Command30.Enabled = True
        End If
        
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If Cancel = 0 And Me.Visible = True Then
       X = VALIDASISEMODIFICO%
   End If
   If Me.Visible = True Then
      Call CIERRARCH("*.*")
      Call BAJALDISCO
      Call FINAL("")
   End If
   
End Sub

Private Sub List1_DblClick()
   List1.Visible = False
End Sub


Private Sub List2_Click()
   List2.ToolTipText = TRIM$(Mid$(List2.TEXT, 25))
End Sub



Private Sub mnRecetaExistente_Click()
   If VALIDASISEMODIFICO% < 0 Then Exit Sub
   NUS% = USNBR% Mod 100
   VENDE% = LOGVEND.VendeVsUsuario(NUS%)
   NRO = SELEC_RECETA(0, 0, , , VENDE%, 1)
   If NRO > 0 Then
     Text5 = ""
     LBLNUEVAFICHATECNICA.Visible = False
     Text5 = NRO
     Call VALORESPORDEFECTO
   End If
   MODIFIC% = 0
   'SIEMPRE QUEDA HABILITADA LA POSIBILIDAD DE AGREGAR UN ARCHIVO
   Command47.Enabled = True
   Command43.Enabled = True
   Command45.Enabled = True
   Combo2.Enabled = True
End Sub

 Function SELEC_RECETA(Optional NRORECETA&, Optional STAT%, Optional NC, Optional SOLOCONCODIGO%, Optional VENDE%, Optional NoVerNoMostrar%) As Long
    'SI NoVerNoMostrar% = 1 NO MUESTRA LAS MARCADAS COMO NO MOSTRAR EN LISTA DE SELECCION
    'SI NoVerNoMostrar% = -1 MUESTRA SOLO LAS MARCADAS COMO NO MOSTRAR EN LISTA DE SELECCION
    VerTodasLasFichasTecnicas% = 0
    If VENDE% > 1 Then 'si vendees > 1 por que es un vendedor y filtra
        'pero si ese usuario tiene derechos para ver todas las fichas tecnicas muestra todas
        If DERACCE%("VETOFITE/NOMESS", "VER TODAS LAS FICHAS TECNICAS") >= 2 Then
           VerTodasLasFichasTecnicas% = 1
        End If
    End If
    Erase RESP_ZELE_VECT
    Campos$ = "N°F.Téc./F9;Fecha/D8;Cod.BB/A12;Raz.Social/A32;Código/A16;Cant./F8.1;ID"
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    FRMZELECHO.Caption = ""
    
    
    Dim obj As New RECORXET
    CAMPOSTABLA$ = "NUMRECETACTP,FECHEMISI,CODBIGBAG,RASOCLI,CODIMASTER,CANTIDAD,RECETACTP_ID"
    TABLA$ = "RECETACTP"
    CONDI$ = "NUMRECETACTP > 0 "
    If STAT% = -1 Then  'SOLO ANULADAS
       CONDI$ = CONDI$ + " AND STATUS = 9"
    ElseIf STAT% = 0 Then ' TODAS MENOS ANULADAS
        'VER EN ESTE CASO TRAE TODAS
'       CONDI$ = CONDI$ + " AND (STATUS = 0 OR STATUS IS NULL)"
    ElseIf STAT% = 2 Then
       CONDI$ = CONDI$ + " AND (STATUS = 1 AND (APROBADA = 0 OR APROBADA IS NULL))"
    End If
    NC = XVALO(NC)
    If NC > 0 Then
       CONDI$ = CONDI$ + " AND NUCLIEN = " & NC
    End If
    '
    If NoVerNoMostrar% > 0 Then
    'SI NoVerNoMostrar% = 1 NO MUESTRA LAS MARCADAS COMO NO MOSTRAR EN LISTA DE SELECCION
    'SI NoVerNoMostrar% = -1 MUESTRA SOLO LAS MARCADAS COMO NO MOSTRAR EN LISTA DE SELECCION
       CONDI$ = CONDI$ + " AND (NoMostEnList  < 1  or NoMostEnList is null)"
    ElseIf NoVerNoMostrar% < 0 Then
       CONDI$ = CONDI$ + " AND (NoMostEnList  = 1 )"
    End If
    
    If SOLOCONCODIGO% > 0 Then
       CONDI$ = CONDI$ + " AND CODIMASTER <> ''"
    ElseIf SOLOCONCODIGO% < 0 Then
       CONDI$ = CONDI$ + " AND (CODIMASTER = '' or CODIMASTER IS NULL)"
    End If
    If VENDE% > 1 And VerTodasLasFichasTecnicas% < 1 Then 'agrego que si no tiene derechos de ver todas filtra, si tiene derecho no filtra
       CONDI$ = CONDI$ + " AND NUVEND = " & VENDE%
    End If
    '
    CONDI$ = CONDI$ + " ORDER BY NUMRECETACTP"
    Set RS1 = obj.RS_GENERAL(TABLA$, CAMPOSTABLA$, CONDI$)
    Call Carga_MSF_Recordset(RS1, Campos$, FRMZELECHO.msf2, 1)
    On Error Resume Next
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    On Error GoTo 0
        
    FRMZELECHO.Width = 12000
50  FRMZELECHO.Show 1
    If XVALO(RESP_ZELE_VECT(7)) > 0 Then ' 7 ES EL ID QUE SI SELECCIONO ES > 0 POR QUE ES LA CONDICION
       SELEC_RECETA = XVALO(RESP_ZELE_VECT(1))
    End If


End Function

 Function NROSOLICITUD(Optional SOLOSINCOTIZAR%, Optional NC) As Long
    Erase RESP_ZELE_VECT
    Campos$ = "ID Sol./f8;N° Vend./F9;Codigo/A16;Descripción/A24;Cod.BB/A20;Cant./f8;Cta./f6;Raz.Social/A32;Ficha Téc./f10"
    
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    FRMZELECHO.Caption = ""
    
    Dim obj As New RECORXET
    CAMPOSTABLA$ = "SOLICOTI_ID,NUVEND,CODIMASTER,DESCRIMASTER, CODBIGBAG,CANTIDAD,NUCLIEN,RASO_CLI,NUMRECETACTP"
    TABLA$ = "SOLICOTI"
    CONDI$ = "SOLICOTI_ID > 0 "
    If SOLOSINCOTIZAR% > 0 Then  'SOLO SIN COTIZAR
       CONDI$ = CONDI$ + " AND (NROCOTIZACION = '' OR NROCOTIZACION IS NULL)"
    End If
    NC = XVALO(NC)
    If NC > 0 Then
       CONDI$ = CONDI$ + " AND NUCLIEN = " & NC
    End If
    '
    '
    CONDI$ = CONDI$ + " ORDER BY SOLICOTI_ID "
    Set RS1 = obj.RS_GENERAL(TABLA$, CAMPOSTABLA$, CONDI$)
    Call Carga_MSF_Recordset(RS1, Campos$, FRMZELECHO.msf2, 1)
    On Error Resume Next
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    On Error GoTo 0
        
    FRMZELECHO.Width = 15705
50  FRMZELECHO.Show 1
    If XVALO(RESP_ZELE_VECT(1)) > 0 Then ' 7 ES EL ID QUE SI SELECCIONO ES > 0 POR QUE ES LA CONDICION
       NROSOLICITUD = XVALO(RESP_ZELE_VECT(1))
    End If


End Function



Private Sub mnuAltaMP_Click()
'    AMASTOK07.Command21_Click
    Call CONECRUN("AMASTO07/ALTACODIGO", "ALTA DE NUEVO CODIGO")
    
End Sub


Private Sub mnuCarpetaImagenes_Click()
    RUTAIMAG.Show 1
End Sub

Sub CARGARUTA()
      '
      CODI$ = TRIM$(TXTCODIGO.TEXT)
      If CODI$ = "" Then
         Call COMUNI("Imposible Continuar !!\Debe Seleccionar Un Código")
         Exit Sub
      End If
      
      CD1.Filter = "(*.*)|*.*|"
      CD1.ShowOpen
      Ruta$ = CD1.FileName
      If TRIM(Ruta$) = "" Then Exit Sub
      '
20    DESCRI$ = TRIM$(InputBox$("Ingrese Texto ", "Guardar Como", CODI$))
      If DESCRI$ = "" Then Exit Sub
      If Len(DESCRI$) > 24 Then
        Call COMUNI("Longitud de Descripción \Supera La Cantidad Máxima de Caracteres (24)")
        GoTo 20
      End If
      On Error Resume Next
      CARPETA_ORIGEN$ = Ruta$
      POS1% = InStrRev(Ruta$, "\")
      POS2% = InStrRev(Ruta$, ".")
      CARPETA_DESTINO$ = Control("RUTAIMAN", "RUTAIMAG") & "\" & TRIM$(DESCRI$) & Mid$(Ruta$, POS2%)
      FileCopy CARPETA_ORIGEN$, CARPETA_DESTINO$
      If Err > 0 Then
       A = 1
      End If
      On Error GoTo 0
      'AGREGAR INFORMACION
      Call REPLA(TX$, DESCRI$, 1, 24)
      Call REPLA(TX$, CARPETA_DESTINO$, 32, 512)
      Combo2.AddItem TX$

End Sub



Private Sub MnuGuardar_Click()
   If VALIDACIONGRABACION% < 0 Then Exit Sub
   ESNUEVO% = 0
   If LBLNUEVAFICHATECNICA.Visible = True Then
      ESNUEVO% = 1
   End If
   '
   SQLX$ = "SELECT * FROM RECETACTP "
   If ESNUEVO% < 1 Then
     SQLX$ = SQLX$ + " WHERE NUMRECETACTP = " & XVALO(Text5)
   Else
     SQLX$ = SQLX$ + " WHERE NUMRECETACTP < 1 "
   End If
   
   On Error GoTo ERROR_GUARDAR
'   FLEXCONN.BeginTrans
   
   Dim RS As ADODB.Recordset
   Set RS = New ADODB.Recordset
   RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   With RS
     If ESNUEVO% > 0 Then
         NEWCOTIZ = NumNuevaReceta
         .AddNew
         !NUMRECETACTP = NEWCOTIZ
         Else
           NEWCOTIZ = XVALO(Text5)
     End If
     
    NC = XVALO(Text1)
    !rasocli = Left$(TXTRAZOCLI, 128)
    !NUCLIEN = NC
    !ESPECIFICLI = Left$(TXTOBSERVA, 1024)
    !CODESPTECTE = Left$(TXTCODESPECIFTECCLTE, 64)
    
    !Cantidad = XVALO(Text2)
    !JUSTA = TRIM$(Left$(TXTJUSTA, 12))
    !NUVEND = XVALO(Text4)
    !FECHEMISI = CVDAT(CVINT(TXTFECHA))
    !CAPKGS = XVALO(TXTCAPACIDAD)
    !CANTIUSOS = Left$(TXTCANTIUSOS, 24)
    !IMPRESION = Left$(TXTIMPRESION, 24)
    !CANTICARAS = XVALO(TXTCANTICARAS)
    !CPOLARGOCM = XVALO(TXTCUERPO_LARGOCM)
    !CPOANCHOCM = XVALO(TXTANCHOCM)
    !CPOALTOCM = XVALO(TXTCUERPOALTOCM)
    !ESLOJOCM = XVALO(TXTESLINGAOJOCM)
    !TELAVENTILADA = SAFETEXT$(TXTTELAVENTILADA)
    !ESLMEDI1CM = XVALO(TXTESLINGAMEDIDA1CM)
    !ESLMED2CM = XVALO(Me.TXTESLINGAMED2EDITABLE)
    !VALVVCDIAMETRO = XVALO(TXTVALVVCDIAMETROCM)
    !VALVVCLARGO = XVALO(TXTVCLARGOCM)
    !VALVVCSLARGOSINO = Left$(SAFETEXT$(TXTVALVVCLARGOCM), 12)
    !VALVVCSLARGO = Left$(Text19, 12)
    !POLLARGO = Left$(TXTPOLLERALARGO, 24)
    !GRMJECUERPO = Left$(TXTGRAMAJECUERPO, 48)
    !GRMJEFONDO = Left$(TXTGRAMAJEFONDO, 48)
    !GRMJETAPA = Left$(TXTGRAMAJETAPA, 48)
    !LAMCPOLAMIN = Left$(TXTLAMCUERPOLAMINADO, 48)
    !LAMFDOLAM = Left$(TXTLAMFONDOLAMINADO, 48)
    !PDOCTIPO = Left$(TXTPDOCPORTADOCUMENTOTIPO, 48)
    !PDOCUBIC = Left$(TXTPDOCPDOCUMENTOUBIC, 48)
    !Compartimentado = Left$(TXTCOMPARTCOMPARTIMENTADO, 48)
    !LINERSINO = Left$(TXTLINERLINER, 48)

    !LINERTIPO = Left$(TXTLINERTIPO, 48)
    !DETAETIQ = Left$(TXTDETALLEETIQUETA, 48)
    !DETAETIQ2 = Left$(TXTDETAETIQUETA2, 48)
    !EMBUFARDO = Left$(TXTEMBALAJEFARDO, 48)
    !TIPOPALLET = Left$(TXTPALLET, 48)
    !UNXPALLET = XVALO(TXTUNPALLET)
    !CODIMASTER = TXTCODIGO
    !CODBIGBAG = TXTCODCONFECCION
    !Producto = Left$(TXTPRODUCTO, 48)
    !ESPIMPRECTE = Left$(Text29, 48)
    !ESPIMIPCTEP = Left$(Text31, 48)
    !Colores = Left$(Text32, 512)
    !CANTCOLORES = XVALO(Text7)
    '!CPOTIPOCOST
    !CPOCOSTPBILO = Left$(TXTCOSTURACONPABILO, 48)
    !CPOCOSTUEXT = Left$(TXTCOSTURAEXTERNA, 48)
    
    !VPDVCGM2 = Left$(TXT_VPDVCGM2, 48)
    !VPDVDGM2 = Left$(TXT_VPDVDGM2, 48)
    !VPDPOLLERAGM2 = Left$(TXT_VPDPOLLERAGM2, 48)
    
    !VPDTELASVALVLAMIN = Left$(TXT_VPDTELASVALVLAMIN, 48)
    !VPDPOLLERA = Left$(TXT_VPDPOLLERA_EDIT, 48)
    !VPDDIVISIONLATERALES = Left$(TXT_VPDDIVISIONLATERALES, 48)
    
    !VPDANCHOUTILMAMPARO = Left$(TXT_VPDANCHOUTILMAMPARO, 48)
    !VPDANCHOUTILMAMPARO = Left$(TXT_VPDANCHOUTILMAMPARO, 48)
    !VPDVDLARGO = Left$(TXT_VPDVDLARGO, 48)
    
    !ESLINGAMEDIDA2CM = Left$(TXTESLINGAMED2CMEDITABLE, 48)
    !VPDVDESCARGACM = Left$(TXT_VPDVDESCARGACM_EDIT, 48)
    !VPDVDESCSEGCM = Left$(TXT_VPDVDESCSEGCM_EDIT, 48)
    
    !VPDLARGOMAMAPAROCM = Left$(TXT_VPDLARGOMAMAPAROCM, 48)
    !VPDDISTANCIABOCAFONDOCM = Left$(TXTDISTBOCAFDOEIDIT, 48)
    !PDFUELLELINERCU = Left$(TXTFUELLELINERCU, 48)
    
    !VPDCargaCm = Left$(TXT_VPDCargaCm_EDIT, 48)
    !VPDvcsegcm = Left$(TXT_VPDvcsegcm_EDIT, 48)
    
    !ESLTIPOCOST = Left$(TXTTIPOCOSTURA, 48)
    !ESLANCHOCM = Left$(TXTESLINGAANCHO, 48)
    !ESLPERIMET = Left$(TXTESLINGAPERIMETRAL, 48)
    !VALVVDDIAM = ROUND(TXTVALVVDDIAMETROCM)
    !VALVVDLARGO = ROUND(TXTVDLARGOCM)
    !VALVVDSLARGO = Left$(TXT_VPDVDSLARGOCM, 48)
    !POLLERALAMIN = Left$(TXTPOLLERALAMINADA, 48)
    !LAMTAPALAM = Left$(TXTTELALAMINADA, 48)
    !CINTLINER = Left$(TXTCINTILLALINER_EDIT, 48)
    !PDOCMEDIDA = Left$(TXTDOCUMENTOUNIDAD, 48)

    !PDOCCANTIDAD = XVALO(TXTPORTADOCUMEDIDA)
    !COMPARTLAMAMP = Left$(Text51, 48)
    !LINERALTO = Left$(Text56, 48)
    !LINERMICRON = SAFETEXT$(ROUND(TXTMICRONAJELINER))
    !OBSGENERAL = Left$(TXTOBSAMPLIA, 1024)
    !VPDTELAMAMPARO = Left$(TXT_VPDTELAMAMPARO, 48)
    !LINERANCHO = SAFETEXT$(XVALO(Text54))
    !NEWVERSIONDE = Left$(TXTNUEVAVERSIONDE, 24)
    !FUELLELINERCU = XVALO(Me.TXTFUELLELINERCU)
    !APROBO = Left$(Me.TXTAPROBO, 32)
    !FEPREPARO = CVDAT(CVINT(Me.TXTFECHPREPARO))
    !Status = 0
    If TRIM$(CotiMasDato.TXTHILACHA) = "" Then CotiMasDato.TXTHILACHA.TEXT = "NO"
    !HILACHA = CotiMasDato.TXTHILACHA.TEXT
    If TRIM$(CotiMasDato.TXTDETEMETAL) = "" Then CotiMasDato.TXTDETEMETAL.TEXT = "Ante evento"
    !DETEMETAL = CotiMasDato.TXTDETEMETAL.TEXT
    'If TRIM$(Me.TXTAPROBO) <> "" Then !FEAPROBO = CVDAT(CVINT(Me.TXTFECHAPROBO))
    !FEAPROBO = Date
    .Update
   End With
   
   If TXTCODIGO.TEXT <> "" Then
      If Command42.Enabled = True Then 'BANDERA PARA SABER SI NO ES LA APROBACION DE FICHA TECNICA.
         CI1% = CODINT(TXTCODIGO)
         CUERPO$ = "Modificación de F.Técnica" + vbCrLf
         Call ANOTAHISMASTER(CI1%, CUERPO$)
         Call CARGAHISTORIAL(CI1%)
         Command4.Enabled = False
      Else
         Command4.Enabled = True
      End If
    End If
    Call GUARDARARCHADJUNTOS(XVALO(Text5))
   
   Text5 = NEWCOTIZ
    
ERROR_GUARDAR:
    If Err.Number <> 0 Then
'       FLEXCONN.RollbackTrans
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description)
       On Error GoTo 0
       
    Else
'      FLEXCONN.CommitTrans
      If NOPRESENTARMENSAJE% < 1 Then
         Call COMUNI("Transacción Realizada")
      End If
      Me.LBLNUEVAFICHATECNICA.Visible = False
      MODIFIC% = 0
    End If


End Sub
Function VALIDACIONGRABACION%()
   VALIDACIONGRABACION% = 1
'   If TRIM$(TXTRAZOCLI) = "" Then
'      Call COMUNI("Falta Ingresar Razon Social")
'      VALIDACIONGRABACION% = -1
'      Call PINTATEXT(TXTRAZOCLI)
'      Exit Function
'   End If
   
   If TRIM$(TXTCODCONFECCION) = "" Then
      Call COMUNI("Falta Ingresar Código de Confección")
      VALIDACIONGRABACION% = -1
      Call PINTATEXT(TXTCODCONFECCION)
      Exit Function
   End If
   
   If XVALO(Text2) < 1 Then
      Call COMUNI("Falta Ingresar Cantidad")
      VALIDACIONGRABACION% = -1
      Call PINTATEXT(Text2)
      Exit Function
   End If
   
   If XVALO(Text4) < 1 Then
      Call COMUNI("Falta Ingresar Vendedor")
      VALIDACIONGRABACION% = -1
      Call PINTATEXT(Text4)
      Exit Function
   End If
   
'8.-En la grabación de la ficha Técnica validar que si Eslinga Perimetral = SI
'el valor de Compart.Dist. Boca y Fdo.(cm)  .  debe ser >= 10 en caso
'que no se cumpla presentar un mensaje alertando esta situación con opción de continuar
'y pasar por alto la advertencia.
   If TRIM$(UCase$(TXTESLINGAPERIMETRAL)) = "SI" Then
      If XVALO(TXTDISTBOCAFDOEIDIT) < 10 Then
         OPX% = COMALTER%("Atención !!!\Eslinga Perimetral = SI \El Valor de Dist. Boca y Fdo. (cm) debe Ser >= 10\\Cancelar Grabación\Grabar")
         If OPX% < 2 Then
'            VALIDACIONGRABACION% = -1
            Call PINTATEXT(TXTDISTBOCAFDOEIDIT)
            Exit Function
         End If
      End If
   End If
   
End Function


Private Sub mnuGuardarComoNuevaFtecnica_Click()
   If XVALO(Text5) > 0 Then
      If VALIDASISEMODIFICO% < 0 Then Exit Sub
      If CantRegistros&("RECETACTP", "NUMRECETACTP = " & XVALO(Text5)) < 1 Then
         Call COMUNI("Debe Seleccionar Una Cotización Que esté Correctamente Guardada en al Base de Datos")
         Exit Sub
      End If
      
      OPX% = COMALTER%("Confirmación de Grabación de Nueva Ficha Técnica en Base a la Ficha Técnica Activa\\Cancelar\Si Grabar")
      If OPX% > 1 Then
         LBLNUEVAFICHATECNICA.Visible = True
         TXTFECHA = FECHATEX$(HOY(HO$))
         Me.TXTCODIGO = ""
         TXTDESCRICOD = ""
         Me.TXTNUEVAVERSIONDE = ""
         Call MnuGuardar_Click
         
         LBLNUEVAFICHATECNICA.Visible = False
         
         MODIFIC% = 0
      End If
   End If
End Sub


Private Sub mnuNuevaReceta_Click()
    Call VALIDASISEMODIFICO
    Call BLANFORMU
    LBLNUEVAFICHATECNICA.Visible = True
    Text5 = SAFETEXT$(NumNuevaReceta)
    TXTFECHA = FECHATEX$(HOY(HO$))
    MODIFIC% = 0
    NUS% = USNBR% Mod 100
    Text4 = LOGVEND.VendeVsUsuario(NUS%)
End Sub
Sub BLANFORMU()
'   Call LIMPIARTEXT(Me, 1, 1)
'ESTO ES POSTERIOR ME DIJERON QUE HABIA CASILLEROS QUE NO SE LIMPIABAN Y PARA NO
'REVISAR UNO POR UNO UTILICE ESTAS 4 FILAS Y MANTUVE EL RESTO
    Call LIMPIARTEXTSEGUNCOLOR(Me, 1, &H80FFFF, &H80FF80)   'amarillo y verde
    Call LIMPIARTEXTSEGUNCOLOR(Me, 1, &H80000002, &HFFFFFF)      'azul y blanco
    Call LIMPIAR_CONTROL_ENCONTENEDOR(Me, Frame2, &HE0E0E0)
    Call LIMPIAR_CONTROL_ENCONTENEDOR(Me, Frame1, &HE0E0E0)

   Call LIMPIARTEXTSEGUNCOLOR(Me, 1, &H80FFFF, &H80FF80)
'   Call LIMPIAR_CONTROL_ENCONTENEDOR(COTICONT, Frame10(4), &HE0E0E0, &HFFFFFF)
   Text1 = ""
   TXTCANTICARAS = ""
    TXT_VPDVDESCARGACM = ""
    TXT_VPDVDESCARGACM_EDIT = ""
    TXTPOLLERALAMINADA = ""
    TXT_VPDVDGM2 = ""
    TXT_VPDPOLLERAGM2 = ""
    TXT_VPDVDDIAMENTRO = ""

    TXTCODIGO = ""
    TXTDESCRICOD = ""
    TXTCODCONFECCION = ""
    TXTDESCRICODCONFECCION = ""
    TXTNUEVAVERSIONDE = ""
    
    TXTNUMCOTIZACION = ""
    TXTREVISION = ""
    Me.TXTOBSAMPLIA = ""
    Me.TXTOBSERVA = ""
    TXTCODESPECIFTECCLTE = ""
    TXTESLINGAMEDIDA2CM = ""
    TXTESLINGAMED2CMEDITABLE = ""
    TXT_VPDFESLINGA = ""
    TXT_VPDVCDIAMETROCM = ""
    TXT_VPDVCLARGOCM = ""
    TXT_VPDVDSLARGOCM = ""
    TXT_VPDVCGM2 = ""
    TXT_VPDTELASVALVLAMIN = ""
    TXT_VPDPOLLERA = ""
    TXT_VPDPOLLERA_EDIT = ""
    TXT_VPDVDLARGO = ""
    TXT_VPDDIVISIONLATERALES = ""
    TXT_VPDANCHOUTILMAMPARO = ""
    TXT_VPDTELAMAMPARO = ""
    TXT_VPDALTOLINER = ""
    TXT_VPDANCHO = ""
    TXT_VPDFUELLELINERCU = ""
    TXTFUELLELINERCU = ""
    TXT_VPDDISTANCIABOCAFONDOCM = ""
    TXTDISTBOCAFDOEIDIT = ""
    TXT_VPDLARGOMAMAPAROCM = ""
    TXT_VPDVDESCSEGCM = ""
    TXT_VPDVDESCSEGCM_EDIT = ""
    TXT_VPDVDDIAMENTRO = ""
    TXT_VPDVDESCARGACM = ""
    TXT_VPDVDESCARGACM_EDIT = ""
    TXTPREPARO = ""
    Me.TXTAPROBO = ""
    TXTFECHPREPARO = ""
    TXTFECHAPROBO = ""
    Text6.TEXT = ""
    List1.Clear
    Text7 = ""
    Text32 = ""
    TXTFECHA = FECHATEX$(HOY(HO$))
    MSF1.Rows = 1: msf2.Rows = 1: MSF3.Rows = 1: MSF4.Rows = 1
    TXTJUSTA = "NO"
    For i% = 0 To FRMCOTI2.Text8.UBound
      FRMCOTI2.Text8(i%) = ""
      FRMCOTI2.Text9(i%) = ""
      FRMCOTI2.Text10(i%) = ""
      FRMCOTI2.Text11(i%) = ""
      FRMCOTI2.Text12(i%) = ""
      FRMCOTI2.Text13(i%) = ""
    Next i%
    For i% = 0 To FRMCOTI2.Text7.UBound
      FRMCOTI2.Text7(i%) = ""
      FRMCOTI2.Text13(i%) = ""
    Next i%
    'AGREGO TAMBIEN PARA LIMPIAR LA INFO DE FRMCOTI2
    Call LIMPIARTEXTSEGUNCOLOR(FRMCOTI2, 1, &HE0E0E0, &HFFFFFF)               'GRIS CLARO y blanco
    CotiMasDato.TXTHILACHA = "NO"
    CotiMasDato.TXTDETEMETAL = "Ante evento"
End Sub
Private Sub mnuTablasAuxiliares_Click()
   If DERACCE%("TABAUXGE", "ACCESO A TABLAS AUXILIARES GENERALES") >= 2 Then
       Call MOSTRARTABLASCONFIGURADAS(Me.Name)
   End If
End Sub

Private Sub Text1_Change()
   NCLIE = XVALO(Text1.TEXT)
   If NCLIE > 0 And NCLIE < 99000 Then
      Me.TXTRAZOCLI.BackColor = &HC0C0C0
      Me.TXTRAZOCLI.TEXT = rasocli$(NCLIE)
      If CantRegistros&("RECETACTP", "NUMRECETACTP = " & XVALO(Text5)) > 0 Then
      End If
   Else
       If NCLIE = 99000 And Me.TXTRAZOCLI.BackColor = &HC0C0C0 Then
         Me.TXTRAZOCLI = ""
        
       Else
         Me.TXTRAZOCLI.BackColor = &H80FFFF
       End If
   End If
   Command4.Enabled = True
   MODIFIC% = 1
   
End Sub

Private Sub Text2_gotfocus()
    Call PINTATEXT(Text2)
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub

Private Sub Text2_LostFocus()
    Call VALORESPORDEFECTO
    Text2 = TRIM$(FORMATNUM$(XVALO(Text2), "f9.0"))
    
End Sub

Private Sub Text29_GotFocus()
   Call PINTATEXT(Text29)

End Sub

Private Sub Text29_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub


Private Sub Text31_GotFocus()
   Call PINTATEXT(Text31)

End Sub

Private Sub Text31_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub


Private Sub Text32_GotFocus()
   Call PINTATEXT(Text32)

End Sub

Private Sub Text32_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub


Private Sub Text4_Change()
  Text3 = ECOARCH$("VENDEDOR", Chr$(XVALO(Text4.TEXT)))
  MODIFIC% = 1
End Sub

Private Sub Text5_Change()
   
'   If VALIDACIONGRABACION% < 0 Then Exit Sub
    Call BLANFORMU
    MODIFIC% = 0
    If XVALO(Text5) > 0 Then
      If CantRegistros&("RECETACTP", "NUMRECETACTP = " & XVALO(Text5)) < 1 Then
         ESNUEVO% = 1
         Exit Sub
      End If
   Else
      Exit Sub
   End If

   SQLX$ = "SELECT * FROM RECETACTP  WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE NUMRECETACTP = " & XVALO(Text5)
   Set RS = READSET(SQLX$)
   With RS
  
        Text1 = XVALO(!NUCLIEN)
        TXTRAZOCLI = SAFETEXT$(!rasocli)
        TXTOBSERVA = SAFETEXT$(!ESPECIFICLI)
        TXTCODESPECIFTECCLTE = SAFETEXT$(!CODESPTECTE)
        Text2 = XVALO(!Cantidad)
        TXTJUSTA = SAFETEXT$(!JUSTA)
        Text4 = XVALO(!NUVEND)
        TXTFECHA = CVDAT(CVINT(!FECHEMISI))
        TXTCAPACIDAD = XVALO(!CAPKGS)
        TXTCANTIUSOS = SAFETEXT$(!CANTIUSOS)
        TXTIMPRESION = SAFETEXT$(!IMPRESION)
        TXTCANTICARAS = XVALO(!CANTICARAS)
        TXTCUERPO_LARGOCM = XVALO(!CPOLARGOCM)
        TXTANCHOCM = XVALO(!CPOANCHOCM)
        TXTCUERPOALTOCM = XVALO(!CPOALTOCM)
        TXTESLINGAOJOCM = XVALO(!ESLOJOCM)
        TXTTELAVENTILADA = SAFETEXT$(!TELAVENTILADA)
        TXTESLINGAMEDIDA1CM = XVALO(!ESLMEDI1CM)
        Me.TXTESLINGAMED2EDITABLE = XVALO(!ESLMED2CM)
        TXTVALVVCDIAMETROCM = XVALO(!VALVVCDIAMETRO)
        TXTVCLARGOCM = XVALO(!VALVVCLARGO)
        TXTVALVVCLARGOCM = SAFETEXT$(!VALVVCSLARGOSINO)
        Text19 = SAFETEXT$(!VALVVCSLARGO)
        TXTPOLLERALARGO = SAFETEXT$(!POLLARGO)
        TXTGRAMAJECUERPO = SAFETEXT$(!GRMJECUERPO)
        TXTGRAMAJEFONDO = SAFETEXT$(!GRMJEFONDO)
        TXTGRAMAJETAPA = SAFETEXT$(!GRMJETAPA)
        TXT_VPDCargaCm_EDIT = SAFETEXT$(!VPDCargaCm)
        TXT_VPDVDESCSEGCM_EDIT = SAFETEXT$(!VPDVDESCSEGCM)
        TXT_VPDPOLLERA_EDIT = SAFETEXT$(!VPDPOLLERA)
        
        TXTLAMCUERPOLAMINADO = SAFETEXT$(!LAMCPOLAMIN)
        TXTLAMFONDOLAMINADO = SAFETEXT$(!LAMFDOLAM)
        TXTPDOCPORTADOCUMENTOTIPO = SAFETEXT$(!PDOCTIPO)
        TXTPDOCPDOCUMENTOUBIC = SAFETEXT$(!PDOCUBIC)
        TXTCOMPARTCOMPARTIMENTADO = SAFETEXT$(!Compartimentado)
        TXTLINERLINER = SAFETEXT$(!LINERSINO)
        TXT_VPDVDESCARGACM_EDIT = SAFETEXT$(!VPDVDESCARGACM)
        TXT_VPDvcsegcm_EDIT = SAFETEXT$(!VPDvcsegcm)
        TXTLINERTIPO = SAFETEXT$(!LINERTIPO)
        TXTDETALLEETIQUETA = SAFETEXT$(!DETAETIQ)
        TXTDETAETIQUETA2 = SAFETEXT$(!DETAETIQ2)
        TXTEMBALAJEFARDO = SAFETEXT$(!EMBUFARDO)
        TXTPALLET = SAFETEXT$(!TIPOPALLET)
        TXTUNPALLET = XVALO(!UNXPALLET)
        TXTCODIGO = TRIM$(SAFETEXT$(!CODIMASTER))
        If TXTCODIGO = "" Then 'SI NO TIENE CODIGO MUESTRA EL DE NUEVAVERSIONDE
                               'SE VA A MOSTRAR SI SE GENERO COMO NUEVA VERSION Y TODAVIA NO POSEE CODIGO
           TXTNUEVAVERSIONDE = SAFETEXT$(!NEWVERSIONDE)
        End If
        TXTCODCONFECCION = SAFETEXT$(!CODBIGBAG)
        TXTPRODUCTO = SAFETEXT$(!Producto)
        Text29 = SAFETEXT$(!ESPIMPRECTE)
        Text31 = SAFETEXT$(!ESPIMIPCTEP)
        Text32 = SAFETEXT$(!Colores)
        Text7 = XVALO(!CANTCOLORES)
        
        TXTCOSTURACONPABILO = SAFETEXT$(!CPOCOSTPBILO)
        TXTCOSTURAEXTERNA = SAFETEXT$(!CPOCOSTUEXT)
        
        TXTTIPOCOSTURA = SAFETEXT$(!ESLTIPOCOST)
        TXTESLINGAANCHO = SAFETEXT$(!ESLANCHOCM)
        TXTESLINGAPERIMETRAL = SAFETEXT$(!ESLPERIMET)
        TXTVALVVDDIAMETROCM = XVALO(!VALVVDDIAM)
        TXTVDLARGOCM = SAFETEXT$(!VALVVDLARGO)
        TXT_VPDVDSLARGOCM = SAFETEXT$(!VALVVDSLARGO)
        TXTPOLLERALAMINADA = SAFETEXT$(!POLLERALAMIN)
        TXTTELALAMINADA = SAFETEXT$(!LAMTAPALAM)
        TXTCINTILLALINER_EDIT = SAFETEXT$(!CINTLINER)
        TXTPORTADOCUMEDIDA = XVALO(!PDOCCANTIDAD)
        
        TXT_VPDVCGM2 = SAFETEXT$(!VPDVCGM2)
        TXT_VPDVDGM2 = SAFETEXT$(!VPDVDGM2)
        TXT_VPDPOLLERAGM2 = SAFETEXT$(!VPDPOLLERAGM2)
        TXT_VPDTELASVALVLAMIN = SAFETEXT$(!VPDTELASVALVLAMIN)
        
        TXTDOCUMENTOUNIDAD = SAFETEXT$(!PDOCMEDIDA)
        Text51 = SAFETEXT$(!COMPARTLAMAMP)
        Text56 = SAFETEXT$(!LINERALTO)
        TXTMICRONAJELINER = SAFETEXT$(!LINERMICRON)
        TXTOBSAMPLIA = SAFETEXT$(!OBSGENERAL)
        Text54 = SAFETEXT$(!LINERANCHO) & " + Fuelles"
        TXTNUEVAVERSIONDE = SAFETEXT$(!NEWVERSIONDE)
        Me.TXTPREPARO = Text3
        Me.TXTAPROBO = SAFETEXT$(!APROBO)
        Me.TXTFECHPREPARO = CVDAT(CVINT(!FEPREPARO)):
        Me.TXTFECHAPROBO = CVDAT(CVINT(!FEAPROBO)): If TXTFECHAPROBO = "01/01/80" Then TXTFECHAPROBO = ""
        Me.TXTDISTBOCAFDOEIDIT = SAFETEXT$(!VPDDISTANCIABOCAFONDOCM)
        Me.TXTFUELLELINERCU = XVALO(!FUELLELINERCU)
        
        CotiMasDato.TXTHILACHA = SAFETEXT$(!HILACHA) ' datos que estan en otro formulario por que no hay mas lugar
        CotiMasDato.TXTDETEMETAL = SAFETEXT$(!DETEMETAL)
   End With
   LBLNUEVAFICHATECNICA.Visible = False
   MODIFIC% = 0
   RS.Close
   Set RS = Nothing
   Call VALORESPORDEFECTO
   Dim CTP As New Conteplast
   Call CTP.CARGARARCHIVOSENCOMBO(XVALO(Text5), Combo2)
   
'   Call GRABARARCHIVOS(XVALO(TEXT1))
'   Call GRABAMODO(XVALO(TEXT1))
   

    
End Sub
Private Sub Text51_Change()

End Sub

Private Sub Text51_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub


Private Sub Text54_GotFocus()
   Call PINTATEXT(Text54)

End Sub

Private Sub Text54_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub


Private Sub Text54_LostFocus()
    Call VALORESPORDEFECTO
    Text54 = TRIM$(FORMATNUM$(XVALO(Text54), "F12.0")) & " + Fuelles"

End Sub

Private Sub Text56_GotFocus()
   Call PINTATEXT(Text56)

End Sub

Private Sub Text56_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub


Private Sub Text56_LostFocus()
    Call VALORESPORDEFECTO
    Text56 = TRIM$(FORMATNUM$(XVALO(Text56), "F12.0"))

End Sub

Private Sub TXTANCHOCM_GotFocus()
    Call PINTATEXT(TXTANCHOCM)

End Sub

Private Sub TXTANCHOCM_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub


Private Sub TXTANCHOCM_LostFocus()
    If TRIM$(TXTCODCONFECCION) = "" Then
       Call COMUNI("Falta Seleccionar Código de Confección")
       Exit Sub
    End If

    Call VALORESPORDEFECTO
    TXTANCHOCM = TRIM$(FORMATNUM$(XVALO(TXTANCHOCM), "F12.0"))
End Sub

Private Sub TXTCANTICARAS_DblClick()
   If Command6.Enabled = True Then TXTCANTICARAS = ""
   
End Sub

Private Sub TXTCAPACIDAD_GotFocus()
    Call PINTATEXT(TXTCAPACIDAD)

End Sub

Private Sub TXTCAPACIDAD_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub

Private Sub TXTCAPACIDAD_LostFocus()
   Call VALORESPORDEFECTO
   TXTCAPACIDAD = TRIM$(FORMATNUM$(XVALO(TXTCAPACIDAD), "F12.0"))
End Sub

Private Sub TXTCODCONFECCION_Change()
   If TRIM$(TXTCODCONFECCION) = "" Then Me.TXTDESCRICODCONFECCION = ""
   MODIFIC% = 1
   If TRIM$(TXTCODIGO) = "" Then
     Call VERIFICARSEGUNBIGBAG(Me.TXTCODCONFECCION)
   End If
   If TRIM$(TXTCODCONFECCION) <> "" Then
     CONDI$ = "CAMPO1 = '" & TRIM$(TXTCODCONFECCION) & "'"
     TXTDESCRICODCONFECCION = VALOBADA("DATASQL", "CAMPO2", CONDI$, "NOMESS")
   End If
End Sub
Sub VERIFICARSEGUNBIGBAG(CODCONFECCION$)
'        If TRIM$(TXTCODIGO) <> "" Then Exit Sub
        
        SQLX$ = "SELECT * FROM CONFIGBB WITH(NOLOCK)"
        SQLX$ = SQLX$ + " Where CODBIGBAG = '" & CODCONFECCION$ & "' AND (MARBORRA < 1 OR MARBORRA IS NULL)"
        Set RS = READSET(SQLX$)
        With RS
           If Not .EOF Then
                DESCRIPCIONBB = SAFETEXT$(!DESCRIPCIONBB)
                CUERPOLAT = SAFETEXT$(!CUERPOLAT) 'CUERPO C.C.
                LATERAL2 = SAFETEXT$(!LATERAL2) 'LATERALES C.C.
                FONDOPANEL = SAFETEXT$(!FONDOPANEL) 'FONDO C.C.
                TAPAPOLLERA = SAFETEXT$(!TAPAPOLLERA) 'TAPA C.C.
                Pollera = SAFETEXT$(!Pollera) 'POLLERA C.C.
                UPANEL = SAFETEXT$(!FONDOPANEL) 'U PANEL C.C.
                VALVCARGA = SAFETEXT$(!VALVCARGA) 'VALV.CARGA
                VALVDESCAR = SAFETEXT$(!VALVDESCAR) 'VALV.DESC.
                CINTILLAPOLLERA = SAFETEXT$(!CINTIPOLLERA) 'CINTILLA POLLERA
                CINTDESC = SAFETEXT$(!CINTDESC) 'CINTILLA DESC
                CINTDESCSEG = SAFETEXT$(!CINTDESCSEG) 'CINTILLA DESC. SEG.
                CINTCARGPOLL = SAFETEXT$(!CINTCARGPOLL) 'CINTILLA CARGA POLL.
                CINTCARGSEG = SAFETEXT$(!CINTCARGSEG) '"CINTILLA CARG SEG.
                CINTLINER = SAFETEXT$(!CINTLINER) 'CINTILLA LINER.
                Mamparo = SAFETEXT$(!Mamparo) 'MAMPARO
                HILO1 = SAFETEXT$(!HILO1) 'HILO
                Eslinga = SAFETEXT$(!Eslinga) 'ESLINGA
                ESLPERIM = SAFETEXT$(!ESLPERIM) 'ESLINGA PERIM
                ETIQUETA = SAFETEXT$(!ETIQUETA) 'ETIQUETA
                SOBRE = SAFETEXT$(!SOBRE) 'SOBRE
                LINER = SAFETEXT$(!LINER) 'LINER
                ENVOLTORIO = SAFETEXT$(!ENVOLTORIO) 'ENVOLTORIO
            End If
        End With
        RS.Close
        Set RS = Nothing

        If TRIM$(UCase$(TAPAPOLLERA)) = "NO" Or XVALO(Pollera) > 0 Then
            TXTVALVVCDIAMETROCM.Locked = True: TXTVALVVCDIAMETROCM = ""
            TXTVCLARGOCM.Locked = True: TXTVCLARGOCM = "" 'BLOQUEADO
            TXT_VPDCargaCm_EDIT.Locked = True: TXT_VPDCargaCm_EDIT = ""
            TXT_VPDvcsegcm_EDIT.Locked = True: TXT_VPDvcsegcm_EDIT = ""
            TXTGRAMAJETAPA.Locked = True: TXTGRAMAJETAPA = ""
            TXT_VPDVCGM2 = "" 'GRIS
            TXTTELALAMINADA.Locked = True: TXTTELALAMINADA = ""
            TXTCOSTURAEXTERNA.Locked = True: TXTCOSTURAEXTERNA = "NO"
            Command36.Enabled = False
            Command11.Enabled = False
            Command28.Enabled = False
            Command21.Enabled = False
         Else
            TXTVALVVCDIAMETROCM.Locked = False
            TXTVCLARGOCM.Locked = False
            TXT_VPDCargaCm_EDIT.Locked = False
            TXT_VPDvcsegcm_EDIT.Locked = False
            TXTGRAMAJETAPA.Locked = False
            TXTTELALAMINADA.Locked = False
            TXTCOSTURAEXTERNA.Locked = False
            Command36.Enabled = True
            Command11.Enabled = True
            Command28.Enabled = True
            Command21.Enabled = True
        End If
        '
        If TRIM$(UCase$(Pollera)) = "NO" Or XVALO(TAPAPOLLERA) > 0 Then
            TXTPOLLERALARGO.Locked = True: TXTPOLLERALARGO = ""
            TXTPOLLERALAMINADA.Locked = True: TXTPOLLERALAMINADA = ""
            TXT_VPDPOLLERA_EDIT.Locked = True: TXT_VPDPOLLERA_EDIT = ""
            TXT_VPDPOLLERAGM2 = ""
            Command38.Enabled = False
            Command27.Enabled = False
        Else
            TXTPOLLERALARGO.Locked = False
            TXTPOLLERALAMINADA.Locked = False
            TXT_VPDPOLLERA_EDIT.Locked = False
            Command38.Enabled = True
            Command27.Enabled = True
        End If
         
        If TRIM$(UCase$(VALVDESCAR)) = "NO" Then
           TXTVALVVDDIAMETROCM.Locked = True: TXTVALVVDDIAMETROCM = ""
           TXTVDLARGOCM.Locked = True: TXTVDLARGOCM = ""
           TXT_VPDVDSLARGOCM = ""
           TXTVDLARGOCM.Locked = True: TXTVDLARGOCM = ""
           TXT_VPDVDGM2.Locked = True: TXT_VPDVDGM2 = ""
           TXT_VPDVDESCARGACM_EDIT.Locked = True: TXT_VPDVDESCARGACM_EDIT = ""
           TXT_VPDVDESCSEGCM_EDIT.Locked = True: TXT_VPDVDESCSEGCM_EDIT = ""
           Command26.Enabled = False
        Else
           TXTVALVVDDIAMETROCM.Locked = False
           TXTVDLARGOCM.Locked = False
           TXTVDLARGOCM.Locked = False
           TXT_VPDVDGM2.Locked = False
           TXT_VPDVDESCARGACM_EDIT.Locked = False
           TXT_VPDVDESCSEGCM_EDIT.Locked = False
           Command26.Enabled = True
        End If

        If TRIM$(UCase$(VALVDESCAR)) = "NO" And TRIM$(UCase$(VALVCARGA)) = "NO" Then
           TXT_VPDTELASVALVLAMIN = ""
        End If
        
        If TRIM$(UCase$(Mamparo)) = "NO" Then
            TXTCOMPARTCOMPARTIMENTADO.Locked = True: TXTCOMPARTCOMPARTIMENTADO = ""
            TXTDISTBOCAFDOEIDIT.Locked = True: TXTDISTBOCAFDOEIDIT = ""
            TXT_VPDLARGOMAMAPAROCM = ""
            TXT_VPDDIVISIONLATERALES = ""
            TXT_VPDANCHOUTILMAMPARO = ""
            TXT_VPDTELAMAMPARO = ""
            Command16.Enabled = False
       Else
            TXTCOMPARTCOMPARTIMENTADO.Locked = False
            TXTDISTBOCAFDOEIDIT.Locked = False
            Command16.Enabled = True
            
       End If
       'ESTO LO TRAJE DE VALORESPORDEFECTO POR QUE SI NO CUANDO TIENE CODIGO NO LO CALCULA  06/05/20
       
       If TRIM$(UCase$(TXTCOMPARTCOMPARTIMENTADO)) = "NO" Or TRIM$(TXTCOMPARTCOMPARTIMENTADO) = "" Then
           TXT_VPDLARGOMAMAPAROCM = ""
       Else
          If IsNumeric(TXTCUERPOALTOCM) Then
              'alto - (distancia de fondo x 2)
              XVALO (TXTCUERPOALTOCM)
'              TXT_VPDLARGOMAMAPAROCM = XVALO(TXTCUERPOALTOCM) - (XVALO(TXTDISTBOCAFDOEIDIT) * 2)
          End If
'              If XVALO(TXTCUERPOALTOCM) <= 130 Then
'                 TXT_VPDLARGOMAMAPAROCM = XVALO(TXTCUERPOALTOCM) - 20
'               Else
'                  If XVALO(TXTCUERPOALTOCM) >= 201 Then
'                     TXT_VPDLARGOMAMAPAROCM = XVALO(TXTCUERPOALTOCM) - 40
'                  Else
'                      TXT_VPDLARGOMAMAPAROCM = XVALO(TXTCUERPOALTOCM) - 30
'                  End If
'              End If
'           End If
        End If

        If TRIM$(TXTCOMPARTCOMPARTIMENTADO) = "NO" Or TRIM$(TXTCOMPARTCOMPARTIMENTADO) = "" Then
            'TXTDISTBOCAFDOEIDIT = ""
            TXT_VPDDISTANCIABOCAFONDOCM = ""
            TXT_VPDLARGOMAMAPAROCM = ""
        Else
            'TXTDISTBOCAFDOEIDIT = (XVALO(TXTCUERPOALTOCM) - XVALO(TXT_VPDLARGOMAMAPAROCM)) / 2
'            TXT_VPDDISTANCIABOCAFONDOCM = (XVALO(TXTCUERPOALTOCM) - XVALO(TXT_VPDLARGOMAMAPAROCM)) / 2
            
'            TXT_VPDLARGOMAMAPAROCM = XVALO(TXTCUERPOALTOCM) - (XVALO(TXTDISTBOCAFDOEIDIT) * 2)
        End If
       ' FIN ESTO LO TRAJE DE VALORESPORDEFECTO POR QUE SI NO CUANDO TIENE CODIGO NO LO CALCULA  06/05/20
       'CAMBIE PARA QUE MODIFIQUE EL NO EDITABLE Y MANTENGA LA INFO QUE SE INGRESA EN EL CASILLERO EDITABLE.
 End Sub

Private Sub TXTCODCONFECCION_DblClick()
   If Command3.Enabled = True Then TXTCODCONFECCION = ""
End Sub

Private Sub txtCodigo_Change()
   TXTDESCRICOD = ""
   If TRIM$(TXTCODIGO) <> "" Then
      CI1% = CODINT%(TXTCODIGO)
      REV$ = REVICOD$(CI1%)
      Me.TXTREVISION = REVICOD$(CI1%)
      TXTDESCRICOD = DESCRIT0$(CI1%)
      Call BLOQUEAR_CONTROLES_ENCONTENEDOR(Me, 2, 1, 1, 1, 0, 1, 0, Me.SSTab1)
      Call BLOQUEAR_CONTROLES_ENCONTENEDOR(Me, 2, 1, 1, 1, 0, 1, 0, Me.Frame1)
      Call BLOQUEAR_CONTROLES_ENCONTENEDOR(Me, 2, 1, 1, 1, 0, 1, 0, Me.Frame2)
      CMD1.Enabled = True 'DEJA HABILITADO EL BOTON DE CANTIDAD JUSTA SOLO SE PUEDE CAMBIAR CON CLAVE
      
      mnuGuardar.Enabled = False
      Command4.Enabled = False
      Command39.Enabled = True
      Call CARGAHISTORIAL(CI1%)
      mnuModificarFTG.Enabled = True
   Else
      Call BLOQUEAR_CONTROLES_ENCONTENEDOR(Me, 1, 1, 1, 1, 0, 1, 0, Me.SSTab1)
      Call BLOQUEAR_CONTROLES_ENCONTENEDOR(Me, 1, 1, 1, 1, 0, 1, 0, Me.Frame1)
      Call BLOQUEAR_CONTROLES_ENCONTENEDOR(Me, 1, 1, 1, 1, 0, 1, 0, Me.Frame2)
      Call VERIFICARSEGUNBIGBAG(Me.TXTCODCONFECCION)
      mnuModificarFTG.Enabled = False
      mnuGuardar.Enabled = True
      Command4.Enabled = True
      Command39.Enabled = False
   End If
   MODIFIC% = 1
   
End Sub

Sub CARGAHISTORIAL(CI1%)
        List1.Clear
        If CI1% < 1 Then Exit Sub
        
        CONDI$ = "CODINT = " & CI1%
        XX$ = VALOBADA("ANOTAMASTER", "HISTORIAL", CONDI$)
        LONGI& = Len(XX$)
        Inicio& = 1
        Do While Inicio& < LONGI&

            FIN& = InStr(Inicio&, XX$, Chr$(10))
            If FIN& > 0 Then
               TTXX$ = Mid$(XX$, Inicio&, FIN& - Inicio&)
               List1.AddItem TTXX$
               Inicio& = FIN& + 1
            Else
               TTXX$ = Mid$(XX$, Inicio&)
               List1.AddItem TTXX$
               Exit Do
            End If
        Loop
        
           ' ESTO ES PARA POSICIONARSE AL FINAL DE LA LISTA
        TPXI& = List1.ListCount - 15
        If TPXI& < 0 Then TPXI& = 0
        List1.TopIndex = TPXI&
        '
        List1.Refresh

End Sub
Private Sub TXTCOMPARTCOMPARTIMENTADO_DblClick()
   If Command16.Enabled = True Then TXTCOMPARTCOMPARTIMENTADO = ""
End Sub

Private Sub TXTCUERPO_LARGOCM_GotFocus()
    Call PINTATEXT(TXTCUERPO_LARGOCM)

End Sub

Private Sub TXTCUERPO_LARGOCM_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub


Private Sub TXTCUERPO_LARGOCM_LostFocus()
    If TRIM$(TXTCODCONFECCION) = "" Then
       Call COMUNI("Falta Seleccionar Código de Confección")
       Exit Sub
    End If

    Call VALORESPORDEFECTO
    TXTCUERPO_LARGOCM = TRIM$(FORMATNUM$(XVALO(TXTCUERPO_LARGOCM), "F12.0"))
End Sub

Private Sub TXTCUERPOALTOCM_GotFocus()
    Call PINTATEXT(TXTCUERPOALTOCM)

End Sub

Private Sub TXTCUERPOALTOCM_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub

Private Sub TXTCUERPOALTOCM_LostFocus()
    If TRIM$(TXTCODCONFECCION) = "" Then
       Call COMUNI("Falta Seleccionar Código de Confección")
       Exit Sub
    End If

    Call VALORESPORDEFECTO
    TXTCUERPOALTOCM = TRIM$(FORMATNUM$(XVALO(TXTCUERPOALTOCM), "F12.0"))
  
End Sub

Private Sub TXTDETALLEETIQUETA_DblClick()
  If Command19.Enabled = True Then TXTDETALLEETIQUETA = ""
End Sub

Private Sub TXTEMBALAJEFARDO_DblClick()
   If Command20.Enabled = True Then TXTEMBALAJEFARDO = ""
End Sub

Private Sub TXTESLINGAMEDIDA1CM_GotFocus()
   Call PINTATEXT(TXTESLINGAMEDIDA1CM)
End Sub

Private Sub TXTESLINGAMEDIDA1CM_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub


Private Sub TXTESLINGAMEDIDA1CM_LostFocus()
    Call VALORESPORDEFECTO
    TXTESLINGAMEDIDA1CM = TRIM$(FORMATNUM$(XVALO(TXTESLINGAMEDIDA1CM), "F12.0"))

End Sub

Private Sub TXTESLINGAOJOCM_GotFocus()
        Call PINTATEXT(TXTESLINGAOJOCM)

End Sub

Private Sub TXTESLINGAOJOCM_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub

Private Sub TXTESLINGAOJOCM_LostFocus()
   Call VALORESPORDEFECTO
   TXTESLINGAOJOCM = TRIM$(FORMATNUM$(XVALO(TXTESLINGAOJOCM), "F12.0"))

End Sub

Private Sub TXTGRAMAJECUERPO_DblClick()
   If Command9.Enabled = True Then TXTGRAMAJECUERPO = ""
End Sub

Private Sub TXTGRAMAJEFONDO_DblClick()
   If Command10.Enabled = True Then TXTGRAMAJEFONDO = ""
End Sub

Private Sub TXTGRAMAJETAPA_DblClick()
  If Command11.Enabled = True Then TXTGRAMAJETAPA = ""
End Sub

Private Sub TXTIMPRESION_DblClick()
  If Command5.Enabled = True Then TXTIMPRESION = ""
End Sub

Private Sub TXTJUSTA_DblClick()
   TXTJUSTA = "NO"
End Sub

Private Sub TXTLAMCUERPOLAMINADO_DblClick()
   If Command13.Enabled = True Then TXTLAMCUERPOLAMINADO = ""
End Sub

Private Sub TXTLAMFONDOLAMINADO_DblClick()
  If Command12.Enabled = True Then TXTLAMFONDOLAMINADO = ""
End Sub

Private Sub TXTLINERLINER_DblClick()
  If Command17.Enabled = True Then TXTLINERLINER = ""
End Sub

Private Sub TXTLINERTIPO_DblClick()
  If Command18.Enabled = True Then TXTLINERTIPO = ""
End Sub



Private Sub TXTOBS_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub


Private Sub TXTOBSERVA_GotFocus()
    Call PINTATEXT(TXTOBSERVA)
End Sub


Private Sub TXTPDOCPDOCUMENTOUBIC_DblClick()
   If Command15.Enabled = True Then TXTPDOCPDOCUMENTOUBIC = ""
End Sub

Private Sub TXTPDOCPORTADOCUMENTOTIPO_DblClick()
  If Command14.Enabled = True Then TXTPDOCPORTADOCUMENTOTIPO = ""
End Sub

Private Sub TXTPOLLERALAMINADA_Change()
    
    MODIFIC% = 1
End Sub

Private Sub TXTPOLLERALARGO_DblClick()
   If Command38.Enabled = True Then TXTPOLLERALARGO = ""
End Sub

Private Sub TXTPRODUCTO_GotFocus()
   Call PINTATEXT(TXTPRODUCTO)

End Sub

Private Sub TXTPRODUCTO_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub


Private Sub TXTRAZOCLI_DblClick()
    If TXTRAZOCLI.BackColor = &HE0E0E0 Then
       TXTRAZOCLI.BackColor = &H80FFFF
       TXTRAZOCLI.Locked = False
       'Call PINTATEXT(TXTRAZOCLI)
       Command1.Enabled = False
       If XVALO(Text1.TEXT) <> 99000 Then Text1.TEXT = ""
    Else
       TXTRAZOCLI.BackColor = &HE0E0E0
       TXTRAZOCLI.Locked = True
       Command1.Enabled = True
    End If
    MODIFIC% = 1
End Sub

Private Sub TXTRAZOCLI_GotFocus()
    Call PINTATEXT(TXTRAZOCLI)
End Sub


Private Sub TXTRAZOCLI_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub


Private Sub TXTVALVVCDIAMETROCM_DblClick()
   If Command36.Enabled = True Then TXTVALVVCDIAMETROCM = ""
End Sub

Private Sub TXTVALVVCDIAMETROCM_LostFocus()
    Call VALORESPORDEFECTO
    TXTVALVVCDIAMETROCM = TRIM$(FORMATNUM$(XVALO(TXTVALVVCDIAMETROCM), "F12.0"))

End Sub

Private Sub TXTVALVVCLARGOCM_DblClick()
  If Command37.Enabled = True Then TXTVALVVCLARGOCM = ""
End Sub

Private Sub TXTVALVVCLARGOCM_LostFocus()
   Call VALORESPORDEFECTO
   If IsNumeric(TXTVALVVCLARGOCM) Then
      TXTVALVVCLARGOCM = TRIM$(FORMATNUM$(XVALO(TXTVALVVCLARGOCM), "F12.0"))
   End If
End Sub

Private Sub TXTVCLARGOCM_GotFocus()
   Call PINTATEXT(TXTVCLARGOCM)

End Sub

Private Sub TXTVCLARGOCM_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub

Private Sub TXTVCLARGOCM_LostFocus()
    Call VALORESPORDEFECTO
    TXTVCLARGOCM = TRIM$(FORMATNUM$(XVALO(TXTVCLARGOCM), "F12.0"))

End Sub

Function V_Desc_Seg_cm_SEGUN_VD_Diámetro_en_cm(VD_Diámetro_en_cm)
   'BUSCA EL VALOR EN LA TABLA CTP.VALV DEL CAMPO3 SEGUN EL CAMPO2
   CONDI$ = " LISTA =   'CTP.VALV'"
   CONDI$ = CONDI$ + " AND CAST(CAMPO2 AS Integer) = " & VD_Diámetro_en_cm
   V_Desc_Seg_cm_SEGUN_VD_Diámetro_en_cm = XVALO(VALOBADA("DATASQL", "CAMPO3", CONDI$, "NOMESS"))
End Function
Function ValorTelaMamparo_SEGUN_AnchoMamparo(VD_AnchoMampaaro)
   'BUSCA EL VALOR EN LA TABLA CTP.mp DEL CAMPO3 SEGUN EL CAMPO2
   CONDI$ = " LISTA =   'CTP.MP'"
   CONDI$ = CONDI$ + " AND CAST(CAMPO2 AS Integer) = " & VD_AnchoMampaaro
   ValorTelaMamparo_SEGUN_AnchoMamparo = XVALO(VALOBADA("DATASQL", "CAMPO3", CONDI$, "NOMESS"))
End Function


Private Sub TXTVDLARGOCM_GotFocus()
   Call PINTATEXT(TXTVDLARGOCM)

End Sub

Private Sub TXTVDLARGOCM_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub

Private Sub TXTVDLARGOCM_LostFocus()
    Call VALORESPORDEFECTO
     TXTVDLARGOCM = TRIM$(FORMATNUM$(XVALO(TXTVDLARGOCM), "F12.0"))
End Sub

Sub ACTUBASEWALDBOTT_ANT_DESCOMPRIME()
   If Control("BASEWALD", "ESTAGRAB") = "SI" Then Exit Sub
On Error GoTo ERROR_GUARDAR
   Call VERIFICARWINRAR
   RUTADROP$ = Control("COTICONT", "ARCDROPB")
'MsgBox "RUTADROP " & RUTADROP$
   NAMEARCHIVO$ = BUSCOARCHIVO$("SGIW14.mdb.zip")
'MsgBox "NOMBRE ARCHIVO " & NAMEARCHIVO$
   RUTADESTINO$ = Control("CONWALDB", "RUTAWALD")
   POSRIGHT% = InStrRev(RUTADESTINO$, "\")
   RUTADESTINO$ = Mid$(RUTADESTINO$, 1, POSRIGHT% - 1) ' le quito de la cadena el archivo por un cambio que hubo y no tener que revisar todo
   If EXISTE(RUTADROP$ & "\" & NAMEARCHIVO$) > 0 Then
'MsgBox "EXISTE ARCHIVO " & NAMEARCHIVO$
        Call GRACONTROL("BASEWALD", "ESTAGRAB", "SI")
        If EXISTE(RUTADESTINO$ & "\" & Mid$(NAMEARCHIVO$, 1, Len(NAMEARCHIVO$) - 4)) > 0 Then
           Kill RUTADESTINO$ & "\" & Mid$(NAMEARCHIVO$, 1, Len(NAMEARCHIVO$) - 4) 'ELIMINA EL ARCHIVO NO ZIPEADO
        End If
'MsgBox "ELIMINA ARCHIVO " & Mid$(NAMEARCHIVO$, 1, Len(NAMEARCHIVO$) - 4)
        FullPathZip = GetShortPath(RUTADROP$ + "\" + NAMEARCHIVO$)
        
        FolderToDescompress = GetShortPath(RUTADESTINO$)
        If WinRar <> "" Then
         Shell WinRar & " e " & FullPathZip & " " & FolderToDescompress, vbHide
'MsgBox "ENTRO A DESCOMPRIMIR "
        Else
         Shell WinZip & " -e " & FullPathZip & " " & FolderToDescompress, vbHide
        End If
'        FileCopy FolderToDescompress & "\" & "SGIW11.MDB", RUTADESTINO$ & "\" & "SGIW11.MDB"
        
'        FileCopy RUTADROP$ & "\" & NAMEARCHIVO$, RUTADESTINO$
        Screen.MousePointer = 11
        Call T_Espera(20)
        Call GRACONTROL("BASEWALD", "ESTAGRAB", "NO")
        If TRIM$(NAMEARCHIVO$) <> "" Then
          Kill RUTADROP$ & "\" & NAMEARCHIVO$
'MsgBox "ELIMINA ARCHIVO " & RUTADROP$ & "\" & NAMEARCHIVO$
          
        End If
        Screen.MousePointer = 1
        
   End If
ERROR_GUARDAR:
        If Err <> 0 Then
'MsgBox "ERROR " & Err.Description
           Call GRACONTROL("BASEWALD", "ESTAGRAB", "NO")
        End If

End Sub

Function BUSCOARCHIVO$(PARTEDELNOMBRE$)
On Error GoTo errsub
'PARTDELNOMBRE NO SE ESTA USANDO ESTA HARCODEADO EL NOMBRE DEL ARCHIVO
RUTADROP$ = TRIM$(Control("COTICONT", "ARCDROPB"))
If RUTADROP$ = "" Then
  Call COMUNI("FALTA CONFIGURAR RUTA DE ARCHIVO DE INTERCAMBIO CON BASE WALDBOTT")
  Exit Function
End If
Set fso = New FileSystemObject
    
    ' Obtiene el directorio
    Set El_Directorio = fso.GetFolder(RUTADROP$ + "")
    ' Lista los ficheros de esta carpeta
    For Each El_Archivo In El_Directorio.Files
       'Añade la ruta
       NOMBREARCH$ = El_Archivo.Name
       '/FechaArchivoEnDropBox$ = El_Archivo.DateLastModified
       
       'If Left$(UCase$(NOMBREARCH$), 14) = "SGIW11.MDB.ZIP" Then ' ANTES PARA BUSCAR ZIPEADO
       If Left$(UCase$(NOMBREARCH$), 10) = "SGIW14.MDB" Then
          BUSCOARCHIVO$ = NOMBREARCH$
          Exit For
       End If
       
    Next El_Archivo
    Set CTP = Nothing
Exit Function

'Error
errsub:
'Error de permiso denegado
If Err.Number = 70 Then
   Resume Next
   
ElseIf Err.Number = 91 Then
   Screen.MousePointer = vbDefault
   Exit Function
Else
   'MsgBox Err.Description, vbCritical
   Exit Function
End If
End Function
