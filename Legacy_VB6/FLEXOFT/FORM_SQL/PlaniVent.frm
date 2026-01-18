VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form PlaniVent 
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FICHA TÉCNICA BOLSAS"
   ClientHeight    =   9750
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   14190
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9750
   ScaleWidth      =   14190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      BackColor       =   &H00E2D3C0&
      Height          =   9860
      Left            =   10080
      TabIndex        =   75
      Top             =   -120
      Width           =   4095
      Begin VB.TextBox Text9 
         Alignment       =   1  'Right Justify
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
         Left            =   2400
         TabIndex        =   154
         Top             =   6720
         Width           =   1000
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
         Index           =   48
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   153
         TabStop         =   0   'False
         Text            =   "COMISION REAL"
         Top             =   6720
         Width           =   2200
      End
      Begin VB.CommandButton Command22 
         Caption         =   "Aprobar"
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
         Left            =   2040
         TabIndex        =   143
         ToolTipText     =   "Grabar Cotización"
         Top             =   9460
         Width           =   915
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Imprimir"
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
         Left            =   3120
         TabIndex        =   118
         ToolTipText     =   "Imprime Cotización"
         Top             =   9460
         Width           =   915
      End
      Begin VB.TextBox TXTFECHACOTI 
         Alignment       =   1  'Right Justify
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
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   116
         TabStop         =   0   'False
         Top             =   550
         Width           =   900
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
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
         Height          =   280
         Index           =   37
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   115
         TabStop         =   0   'False
         Text            =   "Fecha"
         Top             =   550
         Width           =   960
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
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   114
         TabStop         =   0   'False
         Text            =   "P. DE VENTA REAL U$S"
         Top             =   6360
         Width           =   2200
      End
      Begin VB.TextBox TXTPVTAREAL 
         Alignment       =   1  'Right Justify
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
         Left            =   2400
         TabIndex        =   113
         Top             =   6360
         Width           =   1000
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
         Index           =   34
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   112
         TabStop         =   0   'False
         Text            =   "Observaciones:"
         Top             =   7080
         Width           =   1800
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   111
         Top             =   7440
         Width           =   3900
      End
      Begin VB.TextBox TXTLARGOCORTECOT2 
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
         Left            =   3240
         TabIndex        =   109
         Top             =   2040
         Width           =   765
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
         Left            =   2880
         TabIndex        =   108
         ToolTipText     =   "Selección de Cotizaciones Asociadas a la Planilla de Venta Activa"
         Top             =   240
         Visible         =   0   'False
         Width           =   175
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
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
         Height          =   280
         Index           =   24
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   107
         TabStop         =   0   'False
         Text            =   "Cotización"
         Top             =   240
         Width           =   960
      End
      Begin VB.TextBox TXTNROCOTIZACION 
         Alignment       =   1  'Right Justify
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
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   106
         TabStop         =   0   'False
         Top             =   240
         Width           =   900
      End
      Begin VB.CommandButton Command15 
         Caption         =   "<<<<"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   105
         ToolTipText     =   "Cerrar Pestaña Cotización"
         Top             =   9460
         Width           =   795
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Grabar"
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
         Left            =   1080
         TabIndex        =   104
         ToolTipText     =   "Grabar Cotización"
         Top             =   9460
         Width           =   795
      End
      Begin VB.TextBox TXTPRECIOVENTA 
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
         Left            =   2400
         TabIndex        =   103
         Top             =   6000
         Width           =   1000
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
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   102
         TabStop         =   0   'False
         Text            =   "PRECIO DE VENTA  U$S"
         Top             =   6000
         Width           =   2200
      End
      Begin VB.TextBox TXTRODILLOCOT 
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
         Left            =   2400
         TabIndex        =   101
         Top             =   5400
         Width           =   1000
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
         Index           =   33
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   100
         TabStop         =   0   'False
         Text            =   "RODILLOS (Cm.)"
         Top             =   5400
         Width           =   2200
      End
      Begin VB.TextBox TXTCOSTOCONFECCION 
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
         Left            =   2400
         TabIndex        =   99
         Top             =   5040
         Width           =   1000
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
         Index           =   32
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   98
         TabStop         =   0   'False
         Text            =   "COSTO CONFECCIÓN  U$S"
         Top             =   5040
         Width           =   2200
      End
      Begin VB.TextBox TXTCOSTOIMPRESION 
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
         Left            =   2400
         TabIndex        =   97
         Top             =   4680
         Width           =   1000
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
         Index           =   30
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   96
         TabStop         =   0   'False
         Text            =   "COSTO IMPRESIÓN U$S"
         Top             =   4680
         Width           =   2200
      End
      Begin VB.TextBox TXTANCHOPAPEL 
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
         Left            =   2400
         TabIndex        =   95
         Top             =   4320
         Width           =   1000
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
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   94
         TabStop         =   0   'False
         Text            =   "ANCHO PAPEL Cm."
         Top             =   4320
         Width           =   2200
      End
      Begin VB.TextBox TXTPRECIOPAPEL 
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
         Left            =   2400
         TabIndex        =   93
         Top             =   3960
         Width           =   1000
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
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   92
         TabStop         =   0   'False
         Text            =   "PRECIO PAPEL U$S/Kg"
         Top             =   3960
         Width           =   2200
      End
      Begin VB.TextBox TXTPRECIOPP 
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
         Left            =   2400
         TabIndex        =   91
         Top             =   3600
         Width           =   1000
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
         Index           =   27
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   90
         TabStop         =   0   'False
         Text            =   "PRECIO PP U$S/Kg."
         Top             =   3600
         Width           =   2200
      End
      Begin VB.TextBox TXTGRAMAJE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
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
         Left            =   2400
         TabIndex        =   89
         Top             =   3240
         Width           =   1000
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
         Index           =   26
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   88
         TabStop         =   0   'False
         Text            =   "GRAMAJE Gr/m2"
         Top             =   3240
         Width           =   2200
      End
      Begin VB.TextBox TXTPESO 
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
         Left            =   2400
         TabIndex        =   87
         Top             =   2880
         Width           =   1000
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
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   86
         TabStop         =   0   'False
         Text            =   "PESO GRS."
         Top             =   2880
         Width           =   2200
      End
      Begin VB.TextBox TXTCOSTOPAPEL 
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
         Left            =   2400
         TabIndex        =   85
         Top             =   2400
         Width           =   1000
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
         Index           =   23
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   84
         TabStop         =   0   'False
         Text            =   "COSTO PAPEL"
         Top             =   2400
         Width           =   2200
      End
      Begin VB.TextBox TXTLARGOCORTECOT 
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
         Left            =   2400
         TabIndex        =   83
         Top             =   2040
         Width           =   765
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
         Index           =   22
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   82
         TabStop         =   0   'False
         Text            =   "LARGO DE CORTE"
         Top             =   2040
         Width           =   2200
      End
      Begin VB.TextBox TXTPESOPAPEL 
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
         Left            =   2400
         TabIndex        =   81
         Top             =   1680
         Width           =   1000
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
         Index           =   21
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   80
         TabStop         =   0   'False
         Text            =   "PESO PAPEL"
         Top             =   1680
         Width           =   2200
      End
      Begin VB.TextBox TXTGRAMAJEPAPEL 
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
         Left            =   2400
         TabIndex        =   79
         Top             =   1320
         Width           =   1000
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
         Index           =   20
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   78
         TabStop         =   0   'False
         Text            =   "GRAMAJE PAPEL"
         Top             =   1320
         Width           =   2200
      End
      Begin VB.TextBox TXTLARGOCALCULO 
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
         Left            =   2400
         TabIndex        =   77
         Top             =   960
         Width           =   1000
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
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   76
         TabStop         =   0   'False
         Text            =   "LARGO CALCULO"
         Top             =   960
         Width           =   2200
      End
      Begin VB.Label LBLIDSOLICITUDCOTIZACION 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   120
         TabIndex        =   117
         Top             =   600
         Width           =   855
      End
      Begin VB.Label LBLNEWCOTIZACTP 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nueva Cotización"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   120
         TabIndex        =   110
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2D3C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7695
      Index           =   1
      Left            =   0
      TabIndex        =   32
      Top             =   2040
      Width           =   10050
      Begin VB.Frame Frame3 
         BackColor       =   &H00E2D3C0&
         Height          =   850
         Left            =   5640
         TabIndex        =   144
         Top             =   6240
         Visible         =   0   'False
         Width           =   4095
         Begin VB.TextBox TXTLABEL 
            Alignment       =   2  'Center
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
            Height          =   285
            Index           =   44
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   152
            TabStop         =   0   'False
            Text            =   "Tipo"
            Top             =   180
            Width           =   720
         End
         Begin VB.TextBox TXTLABEL 
            Alignment       =   2  'Center
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
            Height          =   285
            Index           =   45
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   151
            TabStop         =   0   'False
            Text            =   "Fuelle"
            Top             =   495
            Width           =   720
         End
         Begin VB.TextBox TXTTIPOLINER 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
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
            Left            =   840
            TabIndex        =   150
            ToolTipText     =   "Son 4 fuelles x bolsa"
            Top             =   180
            Width           =   1080
         End
         Begin VB.TextBox TXTFUELLELINER 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
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
            Left            =   840
            TabIndex        =   149
            ToolTipText     =   "Son 4 fuelles x bolsa"
            Top             =   495
            Width           =   1080
         End
         Begin VB.TextBox TXTLABEL 
            Alignment       =   2  'Center
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
            Height          =   285
            Index           =   46
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   148
            TabStop         =   0   'False
            Text            =   "Ancho"
            Top             =   180
            Width           =   720
         End
         Begin VB.TextBox TXTLABEL 
            Alignment       =   2  'Center
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
            Height          =   285
            Index           =   47
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   147
            TabStop         =   0   'False
            Text            =   "Alto"
            Top             =   495
            Width           =   720
         End
         Begin VB.TextBox TXTANCHOLINER 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
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
            Left            =   2880
            TabIndex        =   146
            ToolTipText     =   "Son 4 fuelles x bolsa"
            Top             =   180
            Width           =   1080
         End
         Begin VB.TextBox TXTALTOLINER 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
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
            Left            =   2880
            TabIndex        =   145
            ToolTipText     =   "Son 4 fuelles x bolsa"
            Top             =   480
            Width           =   1080
         End
      End
      Begin VB.TextBox Text32 
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
         Left            =   6060
         Locked          =   -1  'True
         TabIndex        =   138
         ToolTipText     =   "Doble Click Para Limpiar Colores"
         Top             =   3960
         Width           =   3200
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
         Left            =   5880
         TabIndex        =   136
         Top             =   3960
         Width           =   175
      End
      Begin VB.TextBox Text8 
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
         Left            =   5460
         TabIndex        =   135
         Top             =   3960
         Width           =   420
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
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
         Height          =   330
         Index           =   43
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   134
         TabStop         =   0   'False
         Text            =   "Unid. x Paquete"
         Top             =   4440
         Width           =   1800
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
         Left            =   1920
         TabIndex        =   133
         Top             =   4440
         Width           =   175
      End
      Begin VB.TextBox UNIDPAQU 
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
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   132
         Top             =   4440
         Width           =   1455
      End
      Begin VB.TextBox UNIPALLET 
         Alignment       =   2  'Center
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
         TabIndex        =   131
         ToolTipText     =   "Son 4 fuelles x bolsa"
         Top             =   4440
         Width           =   1440
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
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
         Height          =   330
         Index           =   42
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   130
         TabStop         =   0   'False
         Text            =   "Paquetes. x Pallet"
         Top             =   4440
         Width           =   1800
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
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
         Height          =   330
         Index           =   41
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   129
         TabStop         =   0   'False
         Text            =   "Tipo de Hilo"
         Top             =   3960
         Width           =   1800
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
         Left            =   1920
         TabIndex        =   128
         Top             =   3960
         Width           =   175
      End
      Begin VB.TextBox TIPOHILO 
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
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   127
         Top             =   3960
         Width           =   1455
      End
      Begin VB.TextBox SISDESCA 
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
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   126
         Top             =   3495
         Width           =   2655
      End
      Begin VB.TextBox SisCarga 
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
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   125
         Top             =   3030
         Width           =   2655
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
         Left            =   6360
         TabIndex        =   124
         Top             =   3030
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
         Left            =   6360
         TabIndex        =   123
         Top             =   3495
         Width           =   175
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
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
         Height          =   330
         Index           =   39
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   122
         TabStop         =   0   'False
         Text            =   "Sistema de Descarga"
         Top             =   3495
         Width           =   1800
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
         Index           =   38
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   121
         TabStop         =   0   'False
         Text            =   "Material"
         Top             =   240
         Width           =   1800
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
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
         Height          =   330
         Index           =   12
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   120
         TabStop         =   0   'False
         Text            =   "Medida Liner Cm."
         Top             =   3495
         Width           =   1800
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
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
         Left            =   2175
         TabIndex        =   119
         ToolTipText     =   "Son 4 fuelles x bolsa"
         Top             =   3495
         Width           =   1440
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
         Left            =   6360
         TabIndex        =   72
         Top             =   630
         Width           =   175
      End
      Begin VB.TextBox TXTRODILLO 
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
         Height          =   285
         Left            =   7800
         Locked          =   -1  'True
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   5460
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.TextBox TXTMAXIMOCOLORES 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   1395
         Width           =   420
      End
      Begin VB.TextBox TXTLARGODECORTE2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   5460
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.TextBox TXTLARGODECORTE 
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
         Height          =   285
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   5460
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Left            =   1410
         Locked          =   -1  'True
         TabIndex        =   64
         TabStop         =   0   'False
         Text            =   "Máximo de Colores"
         Top             =   1395
         Width           =   1845
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E2D3C0&
         Caption         =   "Adjuntar Archivos"
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
         Left            =   120
         TabIndex        =   57
         Top             =   4800
         Width           =   9135
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
            Left            =   7080
            TabIndex        =   61
            Top             =   240
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
            Left            =   7965
            TabIndex        =   60
            Top             =   240
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
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   59
            Top             =   240
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
            Left            =   240
            TabIndex        =   58
            Top             =   240
            Width           =   850
         End
      End
      Begin VB.TextBox TXTMEDFUELLE 
         Alignment       =   2  'Center
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
         Left            =   6615
         TabIndex        =   12
         ToolTipText     =   "Son 4 fuelles x bolsa"
         Top             =   2010
         Width           =   2640
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
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
         Height          =   330
         Index           =   17
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   53
         TabStop         =   0   'False
         Text            =   "Medida Fuelle(x1)Cm."
         Top             =   2010
         Width           =   1800
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
         Left            =   1920
         TabIndex        =   13
         Top             =   3030
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
         Left            =   6360
         TabIndex        =   11
         Top             =   1395
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
         Left            =   6360
         TabIndex        =   10
         Top             =   1005
         Width           =   175
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
         Left            =   6360
         TabIndex        =   7
         Top             =   2520
         Width           =   175
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
         Left            =   1920
         TabIndex        =   6
         Top             =   2010
         Width           =   175
      End
      Begin VB.CommandButton Command6 
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
         Left            =   1920
         TabIndex        =   4
         Top             =   1005
         Width           =   175
      End
      Begin VB.CommandButton Command5 
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
         Left            =   1920
         TabIndex        =   3
         Top             =   630
         Width           =   175
      End
      Begin VB.CommandButton Command4 
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
         Left            =   1920
         TabIndex        =   2
         Top             =   240
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
         Index           =   8
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   52
         TabStop         =   0   'False
         Text            =   "Observaciones:"
         Top             =   5520
         Width           =   1800
      End
      Begin VB.TextBox TXTOBSERVA 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   5880
         Width           =   9015
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
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
         Height          =   330
         Index           =   1
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   51
         TabStop         =   0   'False
         Text            =   "Sistema de Carga"
         Top             =   3030
         Width           =   1800
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
         Index           =   2
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   50
         TabStop         =   0   'False
         Text            =   "Ancho total Cm."
         ToolTipText     =   "Se contempla la medida total del ancho incluyendo fuelles"
         Top             =   630
         Width           =   1800
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
         Index           =   4
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   49
         TabStop         =   0   'False
         Text            =   "Impresión"
         Top             =   1005
         Width           =   1800
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
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   48
         TabStop         =   0   'False
         Text            =   "Cant. Colores"
         ToolTipText     =   "hasta  4 colores"
         Top             =   1395
         Width           =   1200
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
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   47
         TabStop         =   0   'False
         Text            =   "Fuelle"
         Top             =   2010
         Width           =   1800
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
         Index           =   9
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   46
         TabStop         =   0   'False
         Text            =   "Microperforado"
         Top             =   2520
         Width           =   1800
      End
      Begin VB.TextBox TXTMATERIALBOLSA 
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
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox TXTANCHOTOTAL 
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
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   630
         Width           =   1455
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
         Index           =   10
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   43
         TabStop         =   0   'False
         Text            =   "Producto"
         Top             =   240
         Width           =   1800
      End
      Begin VB.TextBox TXTIMPRESION 
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
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   1005
         Width           =   1455
      End
      Begin VB.TextBox TXTPRODUCTO 
         Alignment       =   2  'Center
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
         Left            =   6615
         TabIndex        =   8
         ToolTipText     =   "Se debe ingresar el futuro producto interno, en la bolsa ej.:  Arroz, Maiz, etc."
         Top             =   240
         Width           =   2640
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
         Index           =   13
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   41
         TabStop         =   0   'False
         Text            =   "Alto Util Cm."
         ToolTipText     =   "Se contempla desde boca abierta hasta costura de hilo"
         Top             =   630
         Width           =   1800
      End
      Begin VB.TextBox TXTALTOUTIL 
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
         Left            =   6615
         Locked          =   -1  'True
         TabIndex        =   9
         ToolTipText     =   "Se contempla desde boca abierta hasta costura de hilo"
         Top             =   630
         Width           =   1440
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
         Index           =   14
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   40
         TabStop         =   0   'False
         Text            =   "Pleno"
         Top             =   1005
         Width           =   1800
      End
      Begin VB.TextBox TXTPLENO 
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
         Left            =   6615
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   1005
         Width           =   2640
      End
      Begin VB.TextBox TXTCANTCOLORES 
         Alignment       =   2  'Center
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
         Left            =   3840
         TabIndex        =   5
         Top             =   1395
         Width           =   495
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
         Index           =   16
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   38
         TabStop         =   0   'False
         Text            =   "Cant. Caras"
         Top             =   1395
         Width           =   1800
      End
      Begin VB.TextBox TXTCANTCARAS 
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
         Left            =   6615
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   1395
         Width           =   2640
      End
      Begin VB.TextBox TXTFUELLE 
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
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   2010
         Width           =   1455
      End
      Begin VB.TextBox TXTMICROPERFORADO 
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
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   2520
         Width           =   2655
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
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         Text            =   "Liner"
         Top             =   3030
         Width           =   1800
      End
      Begin VB.TextBox TXTLINER 
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
         Left            =   2175
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   3030
         Width           =   1440
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
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   137
         TabStop         =   0   'False
         Text            =   "Colores"
         Top             =   3960
         Width           =   1485
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Rodillo"
         Height          =   285
         Left            =   6840
         TabIndex        =   71
         Top             =   5460
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Largo de corte2"
         Height          =   285
         Left            =   4320
         TabIndex        =   69
         Top             =   5460
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Largo de corte"
         Height          =   285
         Left            =   2040
         TabIndex        =   68
         Top             =   5460
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.TextBox Text18 
      Height          =   495
      Left            =   6480
      TabIndex        =   31
      Text            =   "Text18"
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox Text17 
      Height          =   495
      Left            =   6480
      TabIndex        =   30
      Text            =   "Text17"
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Frame FRAME 
      BackColor       =   &H00E2D3C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   0
      TabIndex        =   15
      Top             =   -120
      Width           =   10050
      Begin VB.TextBox TXTLABEL1 
         Alignment       =   1  'Right Justify
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
         Index           =   31
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   142
         TabStop         =   0   'False
         Text            =   "APROBÓ:"
         Top             =   1200
         Width           =   885
      End
      Begin VB.TextBox TXTAPROBO 
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
         Left            =   3180
         Locked          =   -1  'True
         TabIndex        =   141
         TabStop         =   0   'False
         ToolTipText     =   "La Aprobación Actualiza la Estructura del Item."
         Top             =   1200
         Width           =   1995
      End
      Begin VB.TextBox TXTFECHAPROBO 
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
         Left            =   5940
         Locked          =   -1  'True
         TabIndex        =   140
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1000
      End
      Begin VB.TextBox TXTLABEL1 
         Alignment       =   1  'Right Justify
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
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   139
         TabStop         =   0   'False
         Text            =   "FECHA:"
         Top             =   1200
         Width           =   645
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Cotizar"
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
         TabIndex        =   74
         Top             =   1680
         Width           =   1155
      End
      Begin VB.CommandButton Command39 
         Caption         =   "Solic. Cotiz."
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
         Left            =   7200
         TabIndex        =   73
         Top             =   1680
         Width           =   1155
      End
      Begin VB.TextBox TEXT5 
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
         Left            =   2040
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1905
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
         Index           =   3
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   62
         TabStop         =   0   'False
         Text            =   "Código Bolsa"
         Top             =   1680
         Width           =   1800
      End
      Begin VB.TextBox TXTNUMPLAVENTA 
         Alignment       =   1  'Right Justify
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
         Left            =   9000
         Locked          =   -1  'True
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   1200
         Width           =   900
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
         Index           =   18
         Left            =   7155
         Locked          =   -1  'True
         TabIndex        =   54
         TabStop         =   0   'False
         Text            =   "F. Técnica Bolsas N°"
         Top             =   1200
         Width           =   1800
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
         Index           =   0
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Text            =   "Fecha"
         Top             =   1200
         Width           =   900
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
         Index           =   11
         Left            =   2175
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Text            =   "R.Social"
         Top             =   240
         Width           =   840
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
         Index           =   31
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Text            =   "Cliente"
         Top             =   240
         Width           =   840
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
         Left            =   4950
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         ToolTipText     =   "Doble Click Para Limpiar Casillero"
         Top             =   690
         Width           =   4200
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
         Index           =   15
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Text            =   "Vendedor"
         Top             =   690
         Width           =   900
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
         Left            =   4470
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   690
         Width           =   400
      End
      Begin VB.CommandButton Command3 
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
         Left            =   4260
         TabIndex        =   23
         Top             =   690
         Width           =   175
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
         Left            =   1020
         TabIndex        =   22
         Top             =   1200
         Width           =   175
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
         Left            =   1230
         TabIndex        =   21
         Top             =   1200
         Width           =   975
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
         TabIndex        =   0
         ToolTipText     =   "Ingresar Razon Social Completa ej.: Arcor S.A. , Nestlé S.R.L."
         Top             =   240
         Width           =   6150
      End
      Begin VB.TextBox TXTCAPACIDAD 
         Alignment       =   2  'Center
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
         Left            =   2040
         TabIndex        =   1
         ToolTipText     =   "Ingresar la Capacidad de la Bolsa Kg. Ej. 50 Kg."
         Top             =   690
         Width           =   1035
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
         Index           =   40
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Text            =   "Capacidad en Kgs."
         Top             =   690
         Width           =   1600
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
         TabIndex        =   19
         Top             =   240
         Width           =   175
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
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   240
         Width           =   900
      End
      Begin VB.CommandButton Command2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   9240
         Picture         =   "PlaniVent.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   240
         Width           =   650
      End
      Begin MSComDlg.CommonDialog CD1 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label LBLNUEVAPVTA 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "NUEVA FICHA TÉCNICA BOLSAS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   585
         Left            =   4920
         TabIndex        =   56
         Top             =   1560
         Visible         =   0   'False
         Width           =   2175
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "PlaniVent.frx":014A
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "PlaniVent.frx":037E
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   6480
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   6480
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Menu mnuTransacciones 
      Caption         =   "Transacciones"
      Begin VB.Menu mnuNewPV 
         Caption         =   "Nueva Ficha Técnica de Bolsas"
      End
      Begin VB.Menu mnuAPVE 
         Caption         =   "Abrir Ficha Técnica de Bolsas Existente"
      End
      Begin VB.Menu MnuGuardar 
         Caption         =   "Guardar Ficha Técnica de Bolsas Activa"
      End
      Begin VB.Menu mnuPrintPV 
         Caption         =   "Imprimir  Ficha Técnica de Bolsas Activa"
      End
      Begin VB.Menu S0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAprobarFichBol 
         Caption         =   "Aprobar Ficha Técnica de Bolsas Activa"
      End
      Begin VB.Menu S 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbrirCotExistente 
         Caption         =   "Abrir Cotización Existente"
      End
   End
   Begin VB.Menu mnuCL 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu mnuSolCotPend 
         Caption         =   "Solicitud de Cotizaciones Pendientes"
      End
   End
   Begin VB.Menu mnuVer 
      Caption         =   "Ver"
   End
   Begin VB.Menu mnuUt 
      Caption         =   "Utilidades"
      Begin VB.Menu mnuHabiltCot 
         Caption         =   "ReHabilitar Cotización "
      End
   End
   Begin VB.Menu mnuConfig 
      Caption         =   "Configuraciones"
      Begin VB.Menu mnuCostosCotizacion 
         Caption         =   "Valores de Cotización"
      End
      Begin VB.Menu mnuAdjuntos 
         Caption         =   "Seleccionar Carpeta Donde se Guardaran las Imagenes"
      End
   End
   Begin VB.Menu mnuAccesosDirectos 
      Caption         =   "Accesos Directos"
   End
End
Attribute VB_Name = "PlaniVent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    Public MODIFIC%
    Private Declare Function GetShortPathName Lib "KERNEL32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBuffer As Long) As Long
    Dim WinRar As String, WinZip As String
    Dim BRTEMP As Recordset
    Public NOPRESENTARMENSAJE%

Sub ACTUALIZAVALORES()
    Me.TXTGRAMAJEPAPEL = TRIM$(FORMATNUM$(XVALO(Control("COTICONT", "GRAMAPEL")), "F12.3"))
    Me.TXTPRECIOPP = TRIM$(FORMATNUM$(XVALO(Control("COTICONT", "PRECIOPP")), "F12.3"))
    Me.TXTPRECIOPAPEL = TRIM$(FORMATNUM$(XVALO(Control("COTICONT", "PRECIPAP")), "F12.3"))
    Me.TXTCOSTOIMPRESION = TRIM$(FORMATNUM$(XVALO(Control("COTICONT", "COSTIMPR")), "F12.3"))
    Me.TXTCOSTOCONFECCION = TRIM$(FORMATNUM$(XVALO(Control("COTICONT", "COSTCONF")), "F12.3"))

End Sub

Sub BLANFORMU()
        TEXT1 = ""
        TXTRAZOCLI = ""
        TXTCAPACIDAD = ""
        Text4 = ""
        TXTFECHA = ""
        Text2 = ""
        TXTMATERIALBOLSA = ""
        TXTANCHOTOTAL = ""
        Text5 = ""
        TXTIMPRESION = ""
        Text12 = ""
        TXTFUELLE = ""
        TXTMICROPERFORADO = ""
        Text10 = ""
        Text11 = ""
        TXTPLENO = ""
        TXTCANTCARAS = ""
        TXTLINER = ""
        TXTOBSERVA = ""
        TXTCANTCOLORES = ""
        Me.TXTPRODUCTO = ""
        Me.TXTALTOUTIL = ""
        Me.TXTMEDFUELLE = ""
        TXTMAXIMOCOLORES = ""
        TXTCANTCOLORES = ""
        Text8 = ""
        Text32 = ""
        
        Text9 = ""
        Text10 = ""
        Text11 = ""
        Text12 = ""

End Sub
Sub BLANFORMUCOTI()
        TXTLARGOCALCULO = ""
        TXTGRAMAJEPAPEL = ""
        TXTPESOPAPEL = ""
'        TXTLARGOCORTECOT = ""
'        TXTLARGOCORTECOT2 = ""
        TXTCOSTOPAPEL = ""
        
        TXTPESO = ""
        TXTGRAMAJE = ""
        TXTPRECIOPP = ""
        TXTPRECIOPAPEL = ""
        TXTANCHOPAPEL = ""
        TXTCOSTOIMPRESION = ""
        TXTCOSTOCONFECCION = ""
'        TXTRODILLOCOT = ""
'        TXTNROCOTIZACION = ""
        TXTPRECIOVENTA = ""
        Me.TXTFECHACOTI = ""
        Me.Text7 = ""
        TXTPVTAREAL = ""
        Text9 = ""
End Sub

Sub CALCULOPARACOTIZACION()

    'LARGO CALCULO

    ALTOUTIL = XVALO(TXTALTOUTIL)
    CONDI$ = " CAMPO2 = '" & TXTMATERIALBOLSA & "'"
    
    Resultadox = XVALO(VALOBADA("DATASQL", "CAMPO5", CONDI$, "NOMESS"))
    TXTLARGOCALCULO = TRIM$(FORMATNUM$(ALTOUTIL + Resultadox, "f12.3"))
    TXTLARGOCALCULO = REPLACAR(TXTLARGOCALCULO, ".", ",")
    'Alto Util Cm. +
    'BUSCAR: MATERIAL BOLSA EN TABLA DATASQL LISTA MATEBOLS SEGUN EL TIPO DE MATERIAL BOLSA
    'Y DEVOLVER EL VALOR DE COLUMNA LARGO CAMPO 5

    'PESO
    If TRIM$(TXTMATERIALBOLSA) = "" Then
      TXTPESO = ""
    Else
      'ANCHO TOTAL CM * LARGO CALCULO  * GRAMAJE/ 10000
      TXTPESO = TRIM$(FORMATNUM$(XVALO(TXTANCHOTOTAL) * XVALO(TXTLARGOCALCULO) * XVALO(TXTGRAMAJE) / 10000, "F12.3"))
      TXTPESO = REPLACAR(TXTPESO, ".", ",")
    End If
 

    'ANCHO papel
    If TRIM$(UCase$(TXTMATERIALBOLSA)) = "BOLSA POLIPAPEL" Then
'      buscar: en tabla DATASQL LISTA ANCHOTOT EL Ancho total Cm. EN CAMPO2 Y DEVOLVER  EL VALOR DE CAMPO3
       CONDI$ = " CONVERT(FLOAT, campo2) = " & XVALO(TXTANCHOTOTAL)
       CONDI$ = CONDI$ + " and lista = 'ANCHOTOT'"
       Resultadox = XVALO(VALOBADA("DATASQL", "CAMPO3", CONDI$, "NOMESS"))
       TXTANCHOPAPEL = TRIM$(FORMATNUM$(Resultadox, "f12.0"))
     Else
        TXTANCHOPAPEL = ""
    End If
    
    'PESO PAPEL
    'TXTPESOPAPEL = ANCHO PAPEL  * LARGO CALCULO * GRAMAJE PAPEL /10000
    On Error Resume Next
    TXTPESOPAPEL = TRIM$(FORMATNUM$(XVALO(TXTANCHOPAPEL) * XVALO(TXTLARGOCALCULO) * XVALO(TXTGRAMAJEPAPEL) / 10000, "f12.3"))
    TXTPESOPAPEL = REPLACAR(TXTPESOPAPEL, ".", ",")
    If Err.Number > 0 Then TXTPESOPAPEL = ""
    On Error GoTo 0
    '
    'COSTO PAPEL
    On Error Resume Next
    ' PESO PAPEL * PRECIO PAPEL / 1000 * 2
    TXTCOSTOPAPEL = TRIM$(FORMATNUM$(XVALO(TXTPESOPAPEL) * XVALO(TXTPRECIOPAPEL) / 1000 * 2, "f12.3"))
    TXTCOSTOPAPEL = REPLACAR(TXTCOSTOPAPEL, ".", ",")
    If Err.Number > 0 Then TXTCOSTOPAPEL = ""
    On Error GoTo 0
    
    'PRECIO DE VENTA TXTPRECIOVENTA
    ' If IMPRESION = "NO" Then
'      VALOR=PESO * PRECIO PP/1000*2
'   Else
'      VALOR= PESO*PRECIO PP /1000*2+ COSTO IMPRESION
'   End If
'   VALOR = VALOR + COSTO CONFECCION + COSTO PAPEL
'Else
'   If IMPRESION = "NO" Then
'      VALOR=PESO * PRECIO PP/1000*2
'   Else
'      VALOR=PESO * PRECIO PP/1000*2 + COSTO IMPRESION
'   End If
'   VALOR = VALOR + COSTO CONFECCION
'End If
    If TRIM$(UCase$(TXTMATERIALBOLSA)) = "BOLSA POLIPAPEL" Then
         If TRIM$(UCase$(TXTMATERIALBOLSA)) = "BOLSA POLIPAPEL" Then
           If TRIM(UCase$(TXTIMPRESION)) = "SI" Then
              VALOR1 = XVALO(TXTPESO) * XVALO(TXTPRECIOPP) / 1000 * 2
           Else
              VALOR1 = XVALO(TXTPESO) * XVALO(TXTPRECIOPP) / 1000 * 2 + XVALO(TXTCOSTOIMPRESION)
           End If
        End If
        VALOR1 = VALOR1 + XVALO(TXTCOSTOCONFECCION) + XVALO(TXTCOSTOPAPEL)
    Else
        If TRIM(UCase$(TXTIMPRESION)) = "NO" Then
           VALOR1 = XVALO(TXTPESO) * XVALO(TXTPRECIOPP) / 1000 * 2
        Else
           VALOR1 = XVALO(TXTPESO) * XVALO(TXTPRECIOPP) / 1000 * 2 + XVALO(TXTCOSTOIMPRESION)
        End If
        VALOR1 = VALOR1 + XVALO(TXTCOSTOCONFECCION)
    End If
    Me.TXTPRECIOVENTA = TRIM$(FORMATNUM$(VALOR1, "F12.3"))
    TXTPRECIOVENTA = REPLACAR(TXTPRECIOVENTA, ".", ",")
    

End Sub

Sub GUARDARARCHIVOSADJUNTOS_PVTA(NroPvta&)

   'MARCA DE BORRADO
   Call ACTUREGISTRO("DOCURELPVTA", "MARBORRA", "NROPVTA = " & NroPvta&, 1, REG&)
   '
   'RECORRO EL LIST PARA GRABAR
   For i& = 1 To Combo2.ListCount
      Z$ = Combo2.List(i& - 1)
      DESCRI$ = UCase$(TRIM$(Mid$(Z, 1, 24)))
      RUTAX$ = TRIM$(Mid$(Z, 32, 512))
      SQLX$ = "INSERT INTO  DOCURELPVTA (NROPVTA,CODINT ,Descripcion, Ruta, Marborra)"
      SQLX$ = SQLX$ + "VALUES ( " & NroPvta&
      'SQLX$ = SQLX$ + "," & CODINT(TXTCODIGO) 'CI%
      SQLX$ = SQLX$ + ",0"
      SQLX$ = SQLX$ + ",'" & DESCRI$ & "'"
      SQLX$ = SQLX$ + ",'" & TRIM$(RUTAX$) & "'"
      SQLX$ = SQLX$ + ",0"
      SQLX$ = SQLX$ + ")"
      FLEXCONN.Execute (SQLX$)
   Next i&
   '
   'BORRO LOS REGISTROS MARCADOS CON 1
   Call BORRAREGISTROS("DOCURELPVTA", "MARBORRA = 1 AND NROPVTA = " & NroPvta&, REGAF&, "NOMESS")
   


End Sub

Sub PrintCotiPlanillaVenta()

    Call SETPRINPORT("RESTORE")
    PUERTOIMPRE$ = Control$(IDENTER$, "PRINTSER")
    If PUERTOIMPRE$ = "" Then PUERTOIMPRE$ = Control$("*", "PRINTSER")
    If PUERTOIMPRE$ <> "" Then
       Call SETPRINPORT(PUERTOIMPRE$)
    End If
    '
    RUTAEXCEL$ = Control("COTICONT", "EXCCOTPV")
    If EXISTE(RUTAEXCEL$) < 1 Then
       Call COMUNI("Falta Configurar Ruta de Plantilla Excel de Formulario de Ficha Técnica")
       Exit Sub
    End If
    '

    ANCHO = ROUND(XVALO(TXTANCHOTOTAL))
    ALTO = ROUND(XVALO(TXTALTOUTIL))
    MEDIDA$ = ANCHO & " cm x " & ALTO & " cm"
    With VENTABNEW.MSF
       .Rows = 9: .Cols = 2
       .TextMatrix(1, 1) = TXTFECHACOTI 'nombre cliente
       .TextMatrix(2, 1) = TXTRAZOCLI
       .TextMatrix(3, 1) = ROUND(XVALO(Text2)) 'cantidad
       .TextMatrix(4, 1) = TXTMATERIALBOLSA 'bolsa poli
       .TextMatrix(5, 1) = Text5 'Descripcion
       .TextMatrix(6, 1) = MEDIDA$
       .TextMatrix(7, 1) = TXTPVTAREAL
       .TextMatrix(8, 1) = "N°: " & ROUND(XVALO(TXTNROCOTIZACION))
       
    End With
    'OJO SI SE LLAMA A PANTALLA A VENTABNEW CUANDO SE ACEPTA BORRA LAS FILAS QUE ESTAN VACIAS
'    VENTABNEW.Show 1
        
    Call Param_Print_Excel(RUTAEXCEL$, "COTIZ ENVA", VENTABNEW.MSF, PUERTOIMPRE$, "1", 1, 8, "", 1, 0)                       'imprime en excel la 1,2 Y 3 columna de mi msf en la columna 1 a partir de la fila 11 en el excel (este graba E imprime)
    Unload VENTABNEW

    Call SETPRINPORT("RESTORE")





End Sub


Function VALIDACIONGRABACION%(Optional Grabar%)
   VALIDACIONGRABACION% = 1
   If XVALO(Me.TXTNUMPLAVENTA) < 1 Then
      Call COMUNI("Falta Número de Planilla Venta")
      VALIDACIONGRABACION% = -1
      Call PINTATEXT(TXTNUMPLAVENTA)
      Exit Function
   End If

   If TRIM$(TXTRAZOCLI) = "" Then
      Call COMUNI("Falta Ingresar Razón Social")
      VALIDACIONGRABACION% = -1
      Call PINTATEXT(TXTRAZOCLI)
      Exit Function
   End If

'   If TRIM$(Text5) = "" Then
'      Call COMUNI("Falta Código (Debe Aprobar la Ficha Técnica Para Generar el Mismo")
'      VALIDACIONGRABACION% = -1
'      Call PINTATEXT(Text5)
'      Exit Function
'   End If

   If XVALO(TXTCAPACIDAD) <= 0 Then
      Call COMUNI("Falta Ingresar Capacidad")
      VALIDACIONGRABACION% = -1
      Call PINTATEXT(TXTCAPACIDAD)
      Exit Function
   End If
   
   If XVALO(Me.Text4) < 1 Then
      Call COMUNI("Falta Ingresar Vendedor")
      VALIDACIONGRABACION% = -1
      Call PINTATEXT(Text4)
      Exit Function
   End If
'    If XVALO(Me.Text2) < 1 Then
'      Call COMUNI("Falta Ingresar Cantidad")
'      VALIDACIONGRABACION% = -1
'      Call PINTATEXT(Text2)
'      Exit Function
'   End If
   
   If TRIM$(Me.TXTMATERIALBOLSA) = "" Then
      Call COMUNI("Falta Ingresar Material")
      VALIDACIONGRABACION% = -1
      Call PINTATEXT(TXTMATERIALBOLSA)
      Exit Function
   End If

   If XVALO(Me.TXTANCHOTOTAL) <= 0 Then
      Call COMUNI("Falta Ancho Total")
      VALIDACIONGRABACION% = -1
      Call PINTATEXT(TXTANCHOTOTAL)
      Exit Function
   End If
   
   If TRIM$(Me.TXTIMPRESION) = "" Then
      Call COMUNI("Falta Informar si Lleva Impresión")
      VALIDACIONGRABACION% = -1
      Call PINTATEXT(TXTIMPRESION)
      Exit Function
   End If
  
   If TRIM(Me.TXTPRODUCTO) = "" Then
      Call COMUNI("Falta Informar Producto")
      VALIDACIONGRABACION% = -1
      Call PINTATEXT(TXTPRODUCTO)
      Exit Function
   End If
   
   If XVALO(Me.TXTALTOUTIL) <= 0 Then
      Call COMUNI("Falta Informar Alto Util")
      VALIDACIONGRABACION% = -1
      Call PINTATEXT(TXTALTOUTIL)
      Exit Function
   End If

   If TRIM$(UCase$(Me.TXTIMPRESION)) = "SI" Then
        If XVALO(TXTCANTCOLORES) <= 0 Then
         Call COMUNI("Falta Informar Cantidad de Colores")
         VALIDACIONGRABACION% = -1
         Call PINTATEXT(TXTCANTCOLORES)
         Exit Function
        End If
        If TRIM(Me.TXTPLENO) = "" Then
         Call COMUNI("Falta Informar SI/NO Pleno")
         VALIDACIONGRABACION% = -1
         Call PINTATEXT(TXTPLENO)
         Exit Function
        End If
        If XVALO(Me.TXTCANTCARAS) <= 0 Then
         Call COMUNI("Falta Informar Cantidad de Caras")
         VALIDACIONGRABACION% = -1
         Call PINTATEXT(TXTCANTCARAS)
         Exit Function
        End If
        If XVALO(TXTMAXIMOCOLORES) < XVALO(TXTCANTCOLORES) Then
         Call COMUNI("Cantidad Informada de Colores Supera el Máximo Sugerido")
         VALIDACIONGRABACION% = -1
         Call PINTATEXT(TXTCANTCOLORES)
         Exit Function
        End If
    End If
     
    If TRIM(Me.TXTFUELLE) = "" Then
      Call COMUNI("Falta Informar SI/NO Fuelle")
      VALIDACIONGRABACION% = -1
      Call PINTATEXT(TXTFUELLE)
      Exit Function
    End If
    
    If TRIM(UCase$(Me.TXTFUELLE)) = "SI" Then
       If XVALO(TXTMEDFUELLE) <= 0 Then
         Call COMUNI("Falta Informar Medida Fuelle")
         VALIDACIONGRABACION% = -1
         Call PINTATEXT(TXTMEDFUELLE)
         Exit Function
       End If
    End If
    
    If TRIM(Me.TXTMICROPERFORADO) = "" Then
      Call COMUNI("Falta Informar SI/NO MICROPERFORADO")
      VALIDACIONGRABACION% = -1
      Call PINTATEXT(TXTMICROPERFORADO)
      Exit Function
    End If
    
    If TRIM(Me.TXTLINER) = "" Then
      Call COMUNI("Falta Informar SI/NO LINER")
      VALIDACIONGRABACION% = -1
      Call PINTATEXT(TXTLINER)
      Exit Function
    End If
    If Grabar% > 0 Then
      Call MnuGuardar_Click
    End If

    
End Function

Function VALIDASISEMODIFICO%()
    If XVALO(TXTNUMPLAVENTA) > 0 Then
     If MODIFIC% = 1 Then
          If COMALTER%("Se han Modificado Datos de La Planilla de Venta Activa\\Registrar Cambios\Continuar Sin Grabar") = 1 Then
            Call MnuGuardar_Click
          End If
          MODIFIC% = 0
     End If
    End If
   VALIDASISEMODIFICO% = 0
   MODIFIC% = 0
End Function

Sub GENINDIPRES(Optional INIPROFO&)
   '
   Call ABREPRESVEN
   ULTINUM& = XVALO(VALOBADA("PRESUVEN", "NroPres", "NroPres > 0 ORDER BY NroPres DESC"))
   '
   UNUMINDI& = 0
   URE& = ULTREG&("INDIPRES")
   If URE& > 0 Then
      XX$ = REGLEIDO$("INDIPRES", URE&)
      UNUMINDI& = CVS(Left$(XX$, 4))
   End If
   '
   If UNUMINDI& <> ULTINUM& Then
      Screen.MousePointer = 11
      SQLX$ = " SELECT * FROM PRESUVEN WITH(NOLOCK)"
      SQLX$ = SQLX$ + " ORDER BY NroPres ASC "
      '
      Set LISPRETEMP = New ADODB.Recordset
      'LISPRETEMP.CursorLocation = adUseClient
      Set LISPRETEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))    ', FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
      
      With LISPRETEMP
         On Error Resume Next
         .MoveFirst
         If Err < 1 Then
           jj& = 0: NPXA& = 0: FIN& = ULTINUM&
           REBLA$ = REGBLAN("INDIPRES")
           Do While (.EOF = False)
             NPX& = !NroPres
             NCLIE% = !NroClie
             CIXI% = !CodiInt
             FEX = CVINT(!FECHEMIS)
             XXX$ = REBLA$
             '
             '\\\OT-06-0168-WAR-09/05/06///
             CN1$ = "": MARNOVI$ = ""
             On Error Resume Next
             CR1$ = !Cod_Rubro
             NCR1 = !NumeCRubro
             CA1$ = !Cod_Anio
             '\\\OT-06-0497-RG-21/11/06///
             MARNOVI$ = !MarNoVig
             '\\\OT-06-0497-RG-FIN///
             On Error GoTo 0
             '
             If TRIM$(CR1$) <> "" And XVALO(NCR1) > 0 Then
               NCR2$ = TRIM$(Str$(NCR1))
               While Len(NCR2$) < 4: NCR2$ = "0" + NCR2$: Wend
               CN1$ = CR1$ + "-" + NCR2$ + "-" + CA1$ + "   "
             End If
             '\\\OT-06-0168-WAR-FIN///
                NomCli$ = TRIM$(rasocli$(NCLIE%))
                If NCLIE% < 1 Then 'RASON SOCIAL DEL PROSPECT
                    NUMCLI$ = VALOBADA("PRESUVEN", "CODI_CLI", "NroPres = " & NPX& & "")
                    NomCli$ = VALOBADA("BSEGCOM", "RASO_CLI", "CODI_CLI = '" & NUMCLI$ & "'")
                End If

             YY2$ = TRIM$(AJUSTI$(CN1$ + NomCli$, 30)) + " - "
             If NPX& <> NPXA& Then
                  YY2$ = YY2$ + TRIM$(CODEXT$(CIXI%))
                Else
                  YY2$ = YY2$ + "VARIOS"
             End If
             YY2$ = YY2$ + " - " + FECHATEX$(FEX)
             '\\\OT-06-0497-RG-21/11/06///
             If MARNOVI$ <> "" Then
                YY2$ = "(N/Vig) " + YY2$
             End If
             '\\\OT-06-0497-RG-FIN///
             '
             Call REPLA(XXX$, MKS$(NPX&), 1, 4)
             Call REPLA(XXX$, YY2$, 5, 60)
             If NPX& <> NPXA& Then
               jj& = jj& + 1
               Call TRACE(20, jj&, FIN&)
             End If
             Call GRAREG("INDIPRES", XXX$, jj&)
             '
             NPXA& = NPX&
           .MoveNext
           Loop
         End If
      End With
      LISPRETEMP.Close
      Set LISPRETEMP = Nothing
      '
      Call SETULTREG("INDIPRES", jj&)
      Call CIERRARCH("INDIPRES")
      '
   End If
   '
   Call FRESHECHO("INDIPRES")
   Screen.MousePointer = 1
   
End Sub
'
Sub ACTINDIPRES(NPX&)
   '
   Call ABREPRESVEN
   SQLX$ = " SELECT * FROM PRESUVEN WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE NroPres = " & NPX&
   '
   Dim LISPRETEMP As ADODB.Recordset
   Set LISPRETEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))    ', FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
   '
   With LISPRETEMP
      If Not (.BOF And .EOF) Then
        XXX$ = REGBLAN("INDIPRES")
        jj& = 0: NPXA& = 0: YY2$ = ""
        Do While (.EOF = False)
           NCLIEX = !NroClie
           CIXI% = !CodiInt
           FEX = CVINT(!FECHEMIS)
           '
           CN1$ = "": MARNOVI$ = ""
           On Error Resume Next
           CR1$ = !Cod_Rubro
           NCR1 = !NumeCRubro
           CA1$ = !Cod_Anio
           '\\\OT-06-0497-RG-21/11/06///
           MARNOVI$ = !MarNoVig
           '\\\OT-06-0497-RG-FIN///
           On Error GoTo 0
           '
           If TRIM$(CR1$) <> "" And XVALO(NCR1) > 0 Then
             NCR2$ = TRIM$(Str$(NCR1))
             While Len(NCR2$) < 4: NCR2$ = "0" + NCR2$: Wend
             CN1$ = CR1$ + "-" + NCR2$ + "-" + CA1$ + "   "
           End If
           '\\\OT-06-0168-WAR-FIN///
           NomCli$ = TRIM$(rasocli$(NCLIEX))
           If NCLIEX < 1 Then 'RASON SOCIAL DEL PROSPECT
               CDXCLI$ = VALOBADA("PRESUVEN", "CODI_CLI", "NroPres = " & NPX& & "")
               NomCli$ = VALOBADA("BSEGCOM", "RASO_CLI", "CODI_CLI = '" & CDXCLI$ & "'")
           End If

           YY2$ = TRIM$(AJUSTI$(CN1$ + NomCli$, 30)) + " - "
           If NPX& <> NPXA& Then
                YY2$ = YY2$ + TRIM$(CODEXT$(CIXI%))
              Else
                YY2$ = YY2$ + "VARIOS"
           End If
           YY2$ = YY2$ + " - " + FECHATEX$(FEX)
           '\\\OT-06-0497-RG-21/11/06///
           If MARNOVI$ <> "" Then
              YY2$ = "(N/Vig) " + YY2$
           End If
           NPXA& = NPX&
        .MoveNext
        Loop
        '
        If YY2$ <> "" Then
           Call REPLA(XXX$, MKS$(NPX&), 1, 4)
           Call REPLA(XXX$, YY2$, 5, 60)
           Call FILTERPUTRECORD("INDIPRES", 1, MKS$(NPX&), XXX$)
        End If
        Call CIERRARCH("INDIPRES")
      End If
   End With
   '
End Sub

Sub AGenRep_Click()
    '
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

'
Sub Command1_Click()
    Command1.Enabled = False
    USNU% = USNBR% Mod 100
    VENDESEGUNUSUARIO% = XVALO(Control("LOGVEND", SAFETEXT(USNU%)))
    '
    NCLIE = NCLLICONTEPLAST(VENDESEGUNUSUARIO%)

'    NCLIE = ListaSeleccionClientes("", True)
    If NCLIE < 1 Then Exit Sub
    '
    NC = NCLIE
    TEXT1.TEXT = TRIM$(Str$(NC))
    '
99  Command1.Enabled = True
    '
End Sub

Private Sub Command10_Click()
   X = LISTASELECGENERAL("CTP.CARAS", , , 1)
   Me.TXTCANTCARAS = TRIM$(RESP_ZELE_VECT(2))
   Call VALORESPORDEFECTO
End Sub

Private Sub Command11_Click()
   
   X = LISTASELECGENERAL("TABLASI_NO", , , 1)
   Me.TXTLINER = TRIM$(RESP_ZELE_VECT(2))
   Call VALORESPORDEFECTO

End Sub


Private Sub Command12_Click()
   X = LISTASELECGENERAL("ALTOUTIL", , , 1)
   Me.TXTALTOUTIL = TRIM$(RESP_ZELE_VECT(2))
   Call VALORESPORDEFECTO

End Sub

Private Sub Command13_Click()
    Command13.Enabled = False
    If DERACCE("COTIPLVE", "COTIZAR PLANILLA DE VENTAS ACTIVA") < 2 Then GoTo 99
    'NroPvta& = NROSOLICITUDPVTA(1, NC, FT&)
    NroSolCoti& = NROSOLICITUDPVTA(1, NC, FT&)
    NUMPVTA& = XVALO(RESP_ZELE_VECT(1))
    If NUMPVTA& < 1 Then GoTo 99
    TXTNUMPLAVENTA = NUMPVTA&
    
    FT& = XVALO(TXTNUMPLAVENTA)
    If FT& < 1 Then
       Call COMUNI("Debe Haber Una Ficha Técnica Activa")
       GoTo 99
    End If
    '
    If TRIM$(Text5) = "" Then
      Call COMUNI("Falta Código (Debe Aprobar la Ficha Técnica Para Generar el Mismo")
        GoTo 99
    End If
    
    NC = XVALO(TEXT1)
    If NC < 1 Then
       Call COMUNI("Falta Seleccionar Cliente")
       GoTo 99
    End If
    '
    If NUMPVTA& < 1 Then
      Call COMUNI("No hay una Planilla de Venta Activa")
      GoTo 99
    End If
    '
    If CantRegistros&("P_VTA_BOLSAS", "NROPVTA = " & NUMPVTA&) < 1 Then
        Call COMUNI("Debe Seleccionar Una Planilla de Venta Que esté Correctamente Guardada en al Base de Datos")
        GoTo 99
    End If
    '
    



    Call BLANFORMUCOTI
'   Campos$ = "N° Planilla Vta./f16;N° Vend./F9;Ref-Código/A48;Cta./f6;Raz.Social/A32;ID Sol./f8"
    'VALIDAR SI ESTA APROBADA (DEBE TENER CODIGO)
    TXTNUMPLAVENTA = ""
    LBLNUEVAPVTA.Visible = False
    TXTNUMPLAVENTA = FT&
    Call VALORESPORDEFECTO
    If TRIM$(Text5) <> "" Then
       Call BLOQUEAR_CONTROLES_ENCONTENEDOR(Me, 2, 1, 1, 1, 0, 1, 1, Me.Frame)
       Call BLOQUEAR_CONTROLES_ENCONTENEDOR(Me, 2, 1, 1, 1, 0, 1, 1, Me.Frame2(1))
       Call COMUNI("Ficha Técnica Aprobada, Deshabilitada Para Registrar Cambios")
       Command2.Enabled = True
       Command39.Enabled = False
       Command13.Enabled = False
    Else
       Command39.Enabled = True
       Command13.Enabled = True
    End If
    
    NVEN% = XVALO(RESP_ZELE_VECT(2))
    REFECODI$ = RESP_ZELE_VECT(3)
    NC = XVALO(RESP_ZELE_VECT(4))
    RS$ = RESP_ZELE_VECT(5)
    IDSOL& = XVALO(RESP_ZELE_VECT(6))
    LBLIDSOLICITUDCOTIZACION = IDSOL&
    TXTNROCOTIZACION = NumNuevaCotiza 'NumNuevaCotizaPVTA
    Me.TXTFECHACOTI = FECHATEX$(HOY(HO$))
    LBLNEWCOTIZACTP.Visible = True
    Me.Width = 14280
    
    Call ACTUALIZAVALORES
    On Error Resume Next
    Me.TXTGRAMAJE.SetFocus
    On Error GoTo 0
99  Command13.Enabled = True
   
End Sub
 Function NROCOTIBOLSAS(Optional NC, Optional FT&) As Long
    Erase RESP_ZELE_VECT
'select C.NROCOTIZACION,C.NUCLIEN,C.RASOCLI,C.NUMRECETACTP,S.NUVEND,C.FECHEMISI,C.CODIMASTER, C.NoMostEnListBolas ,C.COTIZACTP_ID
'FROM COTIZACTP C
'INNER JOIN P_VTA_BOLSAS S ON C.NUMRECETACTP = S.NroPvta
'FROM COTIZACTP C  INNER JOIN P_VTA_BOLSAS S ON C.NUMRECETACTP = S.NroPvta
    Call CREARVISTACotizaFtecBolsas

    Campos$ = "Nro.Cotiz./F10 '  1"
    Campos$ = Campos$ & ";Cta./F7"                      '  2"
    Campos$ = Campos$ & ";R.Social/A32"                 '  3"
    Campos$ = Campos$ & ";F.Téc.Bolsas/f16"             '  4"
    Campos$ = Campos$ & ";N° Vend./F9"                  '  5"
    Campos$ = Campos$ & ";Fecha/D8"                     '  6
    Campos$ = Campos$ & ";Código/A48"                   '  7"
    Campos$ = Campos$ & ";ID Sol./f8"                   '  8

    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    FRMZELECHO.Caption = ""
    
    Dim obj As New RECORXET
    
'        SQLX$ = SQLX$ + " select COT.NROCOTIZACION AS COT_NROCOTIZACION  ,COT.NUCLIEN AS COT_NUCLIEN"
'        SQLX$ = SQLX$ + " ,COT.RASOCLI AS COT_RASOCLI, COT.NUMRECETACTP AS COT_NUMRECETACTP"
'        SQLX$ = SQLX$ + " ,P_VTA.NUVEND AS P_VTA_NUVEND"
'        SQLX$ = SQLX$ + " ,COT.FECHEMISI AS COT_FECHEMISI , COT.CODIMASTER AS COT_CODIMASTER"
'        SQLX$ = SQLX$ + " ,COT.NoMostEnListBolas AS COT_NoMostEnListBolas"
'        SQLX$ = SQLX$ + " ,COT.COTIZACTP_ID AS COT_COTIZACTP_ID"
'        SQLX$ = SQLX$ + " FROM COTIZACTP COT"
'        SQLX$ = SQLX$ + " INNER JOIN P_VTA_BOLSAS P_VTA "
'        SQLX$ = SQLX$ + " ON COT.NUMRECETACTP = P_VTA.NroPvta"
    
    
    CAMPOSTABLA$ = "COT_NROCOTIZACION,COT_NUCLIEN,COT_RASOCLI"
    CAMPOSTABLA$ = CAMPOSTABLA$ + ",COT_NUMRECETACTP, P_VTA_NUVEND,COT_FECHEMISI"
    CAMPOSTABLA$ = CAMPOSTABLA$ + ",COT_CODIMASTER, COT_COTIZACTP_ID"
    TABLA$ = "CotizaFtecBolsas"
    CONDI$ = "COT_NROCOTIZACION > 0  "
    NC = XVALO(NC)
    If NC > 0 Then
       CONDI$ = CONDI$ + " AND COT_NUCLIEN = " & NC
    End If
    '
    If FT& > 0 Then
       CONDI$ = CONDI$ + " AND COT.NUMRECETACTP = " & FT&
    End If
    '
    CONDI$ = CONDI$ + " ORDER BY COT_NROCOTIZACION DESC "
    Set RS1 = obj.RS_GENERAL(TABLA$, CAMPOSTABLA$, CONDI$)
    Call Carga_MSF_Recordset(RS1, Campos$, FRMZELECHO.MSF2, 1)
    On Error Resume Next
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    On Error GoTo 0
        
    FRMZELECHO.Width = 15705
50  FRMZELECHO.Show 1
    NROCOTIBOLSAS = XVALO(RESP_ZELE_VECT(1))


End Function

 Function NROSOLICITUDPVTA(Optional SOLOSINCOTIZAR%, Optional NC, Optional FT&) As Long
    Erase RESP_ZELE_VECT
    Campos$ = " F.Téc.Bolsas/f16;N° Vend./F9;Código/A48;Cta./f6;Raz.Social/A32;ID Sol./f8"
    
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    FRMZELECHO.Caption = "Solicitudes de Cotización de Bolsas"
    
    Dim obj As New RECORXET
    CAMPOSTABLA$ = "NROPVTA,NUVEND,REFECLIENTE,NUCLIEN,RASO_CLI,SOLICOTIPVTA_ID"
    TABLA$ = "SOLICOTIPVTA"
    CONDI$ = "SOLICOTIPVTA_ID > 0  "
    If SOLOSINCOTIZAR% > 0 Then  'SOLO SIN COTIZAR
       CONDI$ = CONDI$ + " AND (NROCOTIZACION = '' OR NROCOTIZACION IS NULL)"
    End If
    NC = XVALO(NC)
    If NC > 0 Then
       CONDI$ = CONDI$ + " AND NUCLIEN = " & NC
    End If
    '
    If FT& > 0 Then
       CONDI$ = CONDI$ + " AND NROPVTA = " & FT&
    End If
    '
    CONDI$ = CONDI$ + " ORDER BY SOLICOTIPVTA_ID "
    Set RS1 = obj.RS_GENERAL(TABLA$, CAMPOSTABLA$, CONDI$)
    Call Carga_MSF_Recordset(RS1, Campos$, FRMZELECHO.MSF2, 1)
    On Error Resume Next
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    On Error GoTo 0
        
    FRMZELECHO.Width = 15705
50  FRMZELECHO.Show 1
    If XVALO(RESP_ZELE_VECT(6)) > 0 Then ' 7 ES EL ID QUE SI SELECCIONO ES > 0 POR QUE ES LA CONDICION
       NROSOLICITUDPVTA = XVALO(RESP_ZELE_VECT(6))
    End If


End Function
Private Sub Command14_Click()
        Command14.Enabled = False
        ESNUEVO% = 0
        If LBLNEWCOTIZACTP.Visible = True Then
          ESNUEVO% = 1
        End If
        '
        If XVALO(TXTNROCOTIZACION) < 1 Then
           Call COMUNI("Debe Seleccionar Una Cotización")
           GoTo 99
        End If
        
        If XVALO(TXTPVTAREAL) <= 0 Then
           Call COMUNI("Falta Ingresar Precio de Venta Real")
           GoTo 99
        End If
        If XVALO(Text9) <= 0 Then
           Call COMUNI("Falta Ingresar Comisiòn Real")
        End If
        
        SQLX$ = "SELECT * FROM COTIZACTP "
        If ESNUEVO% < 1 Then
         SQLX$ = SQLX$ + " WHERE NROCOTIZACION  = " & XVALO(TXTNROCOTIZACION)
        Else
         SQLX$ = SQLX$ + " WHERE NROCOTIZACION  < 1 "
        End If
        On Error GoTo ERROR_GUARDAR
        FLEXCONN.BeginTrans

       'GUARDA EN TABLA CABEZERA COTIZACTP (MISMA TABLA CABEZERA DE LOS BOLOSONES)
        Dim RS As ADODB.Recordset
        Set RS = New ADODB.Recordset
        RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        With RS
         If ESNUEVO% > 0 Then
             .AddNew
             NEWCOTIZ = NumNuevaCotiza
             !NROCOTIZACION = NEWCOTIZ
             !FECHEMISI = CVDAT(HOY(HO$))
         Else
             NEWCOTIZ = XVALO(TXTNROCOTIZACION)
         End If
         !NROCOTIZACION = NEWCOTIZ
         !rasocli = TXTRAZOCLI
         !Status = 0
         !NUMRECETACTP = XVALO(TXTNUMPLAVENTA)
         !NUCLIEN = XVALO(TEXT1)
         !CODIMASTER = Text5
         !NoMostEnList = 1
         !NoMostEnListBolas = 0
         .Update
        End With
        RS.Close
        Set RS = Nothing
        '
        
        'ACTUALIZA LA TABLA COTIPVTA (DETALLE DE COTIZACION)
        SQLX$ = "SELECT * FROM COTIPVTA "
        If ESNUEVO% < 1 Then
         SQLX$ = SQLX$ + " WHERE NROCOTIZACION  = " & XVALO(TXTNROCOTIZACION)
        Else
         SQLX$ = SQLX$ + " WHERE NROCOTIZACION  < 1 "
         
        End If
        'GUARDA EN TABLA CABEZERA COTIZACTP
        Set RS = New ADODB.Recordset
        RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        With RS
         If ESNUEVO% > 0 Then
             NEWCOTIZ = NEWCOTIZ 'YA LO ASIGNO EN LA CONSULTA DE ARRIBA
             .AddNew
             !NROCOTIZACION = NEWCOTIZ
             !FECHEMISI = CVDAT(HOY(HO$))
             Else
               NEWCOTIZ = NEWCOTIZ
         End If
         !GRAMAJE = TXTGRAMAJE
         !NROCOTIZACION = XVALO(TXTNROCOTIZACION)
          NC = XVALO(TEXT1)
         !raso_cli = Left$(TXTRAZOCLI, 128)
         !NUCLIEN = NC
         !OBSERVACIONES = Left$(Text7, 1024)
         !Cantidad = XVALO(Text2)
         !NUVEND = XVALO(Text4)
         !FECHEMISI = CVDAT(CVINT(TXTFECHA))
'         !Status = 0
         !REFECLIENTE = Left(Me.Text5, 64)
         !PRECIO_PP = FormatNumber(XVALO(Me.TXTPRECIOPP), 3)
         !PRECIOPAPEL = FormatNumber(XVALO(Me.TXTPRECIOPAPEL), 3)
         !ANCHOPAPEL = FormatNumber(XVALO(Me.TXTANCHOPAPEL), 3)
         !COSTOIMPRE = FormatNumber(XVALO(Me.TXTCOSTOIMPRESION), 3)
         !COSTOCONFE = FormatNumber(XVALO(Me.TXTCOSTOCONFECCION), 3)
         !RODILLOS = FormatNumber(XVALO(Me.TXTRODILLO), 3)
         !PRECIOVENTA = FormatNumber(XVALO(Me.TXTPRECIOVENTA), 3)
         !PREVTAREAL = FormatNumber(XVALO(Me.TXTPVTAREAL), 3)
         !Peso = FormatNumber(XVALO(Me.TXTPESO), 3)
         !GRAMAJEPAPEL = FormatNumber(XVALO(Me.TXTGRAMAJEPAPEL), 3)
         !PESOPAPEL = FormatNumber(XVALO(Me.TXTPESOPAPEL), 3)
         !LARGOCORTE = FormatNumber(XVALO(Me.TXTLARGODECORTE), 3)
         !LARGOCORTE2 = FormatNumber(XVALO(TXTLARGOCORTECOT2), 3)
         !COSTOPAPEL = FormatNumber(XVALO(Me.TXTCOSTOPAPEL), 3)
         !LARGOCAL = FormatNumber(XVALO(Me.TXTLARGOCALCULO), 3)
         !NroPvta = XVALO(Me.TXTNUMPLAVENTA)
         !ComPorReal = XVALO(Me.Text9)
         
         .Update
       End With
       'ACTUALIZO LA TABLA DE SOLICITUD DE COTIZACIONES (SOLICOTIPVTA) PARA QUE NO QUEDE COMO PENDIENTES
       CONDI$ = "SOLICOTIPVTA_ID =  " & XVALO(LBLIDSOLICITUDCOTIZACION)
       Call ACTUREGISTRO("SOLICOTIPVTA", "NROCOTIZACION", CONDI$, NEWCOTIZ, REGAF&, "NOMESS")
       Call ACTUREGISTRO("SOLICOTIPVTA", "COTIZAR", CONDI$, 3, REGAF&, "NOMESS") '3 = COTIZADA
       

ERROR_GUARDAR:
    If Err.Number <> 0 Then
       FLEXCONN.RollbackTrans
       
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description)
       On Error GoTo 0
       
    Else
      FLEXCONN.CommitTrans
      If NOPRESENTARMENSAJE% < 1 Then
         Call COMUNI("Transacción Realizada")
      End If
      Me.LBLNEWCOTIZACTP.Visible = False
      MODIFIC% = 0
      RS.Close
      Set RS = Nothing
      Command15_Click
      Call BLANFORMUCOTI
      TXTNROCOTIZACION = ""
    End If

99 Command14.Enabled = True
End Sub

Sub Command15_Click()
 Me.Width = 10170
End Sub

Private Sub Command17_Click()
   Call PrintCotiPlanillaVenta
End Sub

Private Sub Command18_Click()
   X = LISTASELECGENERAL("SISDESCA", , , 1)
   Me.SISDESCA = TRIM$(RESP_ZELE_VECT(2))
   Call VALORESPORDEFECTO
End Sub

Private Sub Command19_Click()
   X = LISTASELECGENERAL("SisCarga", , , 1)
   Me.SisCarga = TRIM$(RESP_ZELE_VECT(2))
   Call VALORESPORDEFECTO

End Sub

Private Sub Command2_Click()
    '
    Call CIERRARCH("*.*")
    Unload Me
    '
End Sub

Private Sub Command20_Click()
   X = LISTASELECGENERAL("TIPOHILO", , , 1)
   Me.TIPOHILO = TRIM$(RESP_ZELE_VECT(2))
   Call VALORESPORDEFECTO
End Sub

Private Sub Command21_Click()
   X = LISTASELECGENERAL("UNIDPAQU", , , 1)
   Me.UNIDPAQU = TRIM$(RESP_ZELE_VECT(2))
   Call VALORESPORDEFECTO

End Sub

Private Sub Command22_Click()
    Command22.Enabled = False
    If FIRMAELECTRONICAOK% <= 0 Then GoTo 99

    NC = XVALO(TEXT1)
    
    NROCOT& = XVALO(TXTNROCOTIZACION)
    Codigo$ = Text5
    CI1% = CODINT%(Codigo$)
    NROFICHATEC& = XVALO(TXTNUMPLAVENTA)
    PREUNIT = XVALO(TXTPVTAREAL)
    PorComisReal# = XVALO(Text9)
    If NC < 1 Then
         Call COMUNI("Falta Ingresar Número de Cliente")
         GoTo 99
    End If
    
    If NROCOT& < 1 Then
         Call COMUNI("Falta Número de Cotización")
         GoTo 99
    End If
    
    If CI1% < 1 Then
         Call COMUNI("Código No Válido o Inexistente")
         GoTo 99
    End If
    
    If NROFICHATEC& < 1 Then
         Call COMUNI("Falta Número de Ficha Técnica")
         GoTo 99
    End If
    
    If PREUNIT <= 0 Then
         Call COMUNI("Falta Precio de Venta Real")
         GoTo 99
    End If
    If PorComisReal# <= 0 Then
         Call COMUNI("Falta Porcentaje de Comisiòn Real")
    End If
    
    CONDI$ = "NROCOTIZACION = " & XVALO(TXTNROCOTIZACION)
    If CantRegistros&("COTIPVTA", CONDI$, "NOMES") < 1 Then
       Call COMUNI("La Cotización Activa no ha Sido Grabada\Deberá Grabar la Misma para Luego Aprobarla")
       GoTo 99
    End If
    '
    On Error GoTo ERROR_GUARDAR
    FLEXCONN.BeginTrans
    Call CREACAMPO("", "COTIPVTA", "STATUS", 3, 0)
    SQLX$ = "UPDATE COTIPVTA SET STATUS = 3 WHERE NROCOTIZACION = " & NROCOT&
    FLEXCONN.Execute (SQLX$)
    
    Call FRMCOTI2.GRABARENPRESUVEN(NC, NROCOT&, 1, CI1%, NVEND%, Codigo$, "", NROFICHATEC&, 2, XVALO(TXTPVTAREAL), Text7, OKX%, 1, PorComisReal#)

ERROR_GUARDAR:
        If Err <> 0 Then
           FLEXCONN.RollbackTrans
           MsgBox Err.Description
           On Error GoTo 0
           Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION")
           Exit Sub
         Else
           FLEXCONN.CommitTrans
           Call COMUNI("Grabación Realizado Exitosamente")
         End If
         
99  Command22.Enabled = True
End Sub

Sub Command3_Click()
    '
    Call SELECHO("VENDEDOR")
    VDEVU$ = TRIM$(VALACT1$("VENDEDOR"))
    If VDEVU$ <> "" Then
       Text4.TEXT = VDEVU$
    End If

    '
End Sub

Private Sub Command39_Click()
   Command39.Enabled = False
   NroPvta& = XVALO(TXTNUMPLAVENTA)
   If NroPvta& < 1 Then
     Call COMUNI("No hay una Planilla de Venta Activa")
     GoTo 99
   End If
   
   If CantRegistros&("P_VTA_BOLSAS", "NROPVTA = " & NroPvta&) < 1 Then
     Call COMUNI("Debe Haber uan Planilla de Ventas Grabada Correctamente\Guarde la Misma y Vuelva a Intenar")
     GoTo 99
   End If
   NOPRESENTARMENSAJE% = 1
   If VALIDACIONGRABACION%(1) < 0 Then GoTo 99  'valida datos y graba
   NOPRESENTARMENSAJE% = 0

   '
   'AVISA SI YA TIENE UNA SOLICITUD PENDIENTE DEL MISMO VENDEDOR
   CONDI$ = "NUVEND = " & XVALO(Text4)
   CONDI$ = CONDI$ + " and NROPVTA = " & NroPvta&
   CONDI$ = CONDI$ + " and COTIZAR = 1"
   If CantRegistros&("SOLICOTIPVTA", CONDI$, "NOMESS") > 0 Then
      OPX% = COMALTER%("Se Ha Detectado una Solicitud de Cotización De esta Planilla de Ventas aún Pendiente\\Cancelar\Generar Nueva Solicitud")
      If OPX% < 2 Then GoTo 99
   End If
   
   On Error GoTo ERROR_GUARDAR
   FLEXCONN.BeginTrans
   
    SQLX$ = "INSERT INTO  SOLICOTIPVTA (NUVEND ,CANTIDAD,  REFECLIENTE, COTIZAR, NROPVTA, NROCOTIZACION  "
    SQLX$ = SQLX$ + ", NUCLIEN, CodiEmpr, MARBORRA, FECHEMISI,RASO_CLI,DESCRIMASTER)"
    SQLX$ = SQLX$ + "VALUES ( " & XVALO(Text4) 'NUVEND
    SQLX$ = SQLX$ + "," & XVALO(Text2) 'CANTIDAD
    SQLX$ = SQLX$ + ",'" & TRIM$(Text5) & "'" 'CODIGO
    SQLX$ = SQLX$ + ",1" 'COTIZAR 1(PENDIENTE DE COTIZAR) '3 = COTIZADA
    SQLX$ = SQLX$ + " ," & NroPvta 'nrocotizacion
    SQLX$ = SQLX$ + " ,0" 'NROCOTIZACION
    SQLX$ = SQLX$ + " ," & XVALO(TEXT1) 'NUCLIEN
    SQLX$ = SQLX$ + " ,0" 'CodiEmpr
    SQLX$ = SQLX$ + " ,0" 'MARBORRA
    SQLX$ = SQLX$ + " ,'" & FETEXCOR1$(HOY(HO$)) & "'" 'FECHEMISI
    SQLX$ = SQLX$ + " ,'" & TRIM$(Left$(TXTRAZOCLI, 64)) & "'" 'RASO_CLI
    SQLX$ = SQLX$ + " ,''" '
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
   X = LISTASELECGENERAL("MATEBOLS", , , 1)
   Me.TXTMATERIALBOLSA = TRIM$(RESP_ZELE_VECT(2))
   Call VALORESPORDEFECTO

End Sub




Private Sub Command43_Click()
    Command43.Enabled = False
    NUMPVTA& = XVALO(TXTNUMPLAVENTA)
    If NUMPVTA& < 1 Then
          Call COMUNI("Debe Haber una Planilla de Venta Activa")
          Call PINTATEXT(TXTNUMPLAVENTA)
          GoTo 99
    End If
      
    If CantRegistros&("P_VTA_BOLSAS", "NROPVTA = " & NUMPVTA&) < 1 Then
        Call COMUNI("Debe Seleccionar Una Planilla de Venta Que esté Correctamente Guardada en al Base de Datos")
        GoTo 99
    End If
    '
    If Combo2.ListIndex >= 0 Then
      OPX% = COMALTER%("Confirmación de Eliminación Archivo Adjunto Seleccionado\\Cancelar\Si Eliminar")
      If OPX% = 2 Then
       Combo2.RemoveItem (Combo2.ListIndex)
       Call GUARDARARCHADJUNTOS_PVTA(NUMPVTA&)
      End If
    Else
       Call COMUNI("Atención No Hay Ningún Archivo Seleccionado")
       GoTo 99
    End If
    
99  Command43.Enabled = True


End Sub
Sub GUARDARARCHADJUNTOS_PVTA(NUMPVTA&)
   'MARCA DE BORRADO
   Call ACTUREGISTRO("DOCURELPVTA", "MARBORRA", "NROPVTA = " & NUMPVTA&, 1, REG&)
   '
   'RECORRO EL LIST PARA GRABAR
   For i& = 1 To Combo2.ListCount
      Z$ = Combo2.List(i& - 1)
      DESCRI$ = UCase$(TRIM$(Mid$(Z, 1, 24)))
      RUTAX$ = TRIM$(Mid$(Z, 32, 512))
      SQLX$ = "INSERT INTO DOCURELPVTA (NROPVTA,CODINT ,Descripcion, Ruta, Marborra)"
      SQLX$ = SQLX$ + "VALUES ( " & NUMPVTA&
      'SQLX$ = SQLX$ + "," & CODINT(TXTCODIGO) 'CI%
      SQLX$ = SQLX$ + ",0"
      SQLX$ = SQLX$ + ",'" & DESCRI$ & "'"
      SQLX$ = SQLX$ + ",'" & TRIM$(RUTAX$) & "'"
      SQLX$ = SQLX$ + ",0"
      SQLX$ = SQLX$ + ")"
      FLEXCONN.Execute (SQLX$)
   Next i&
   '
   'BORRO LOS REGISTROS MARCADOS CON 1
   Call BORRAREGISTROS("DOCURELPVTA", "MARBORRA = 1 AND NROPVTA = " & NUMPVTA&, REGAF&, "NOMESS")
   


End Sub
Private Sub Command47_Click()
    Command47.Enabled = False
    NUMPVTA& = XVALO(Me.TXTNUMPLAVENTA)
    If NUMPVTA& < 1 Then
          Call COMUNI("Debe Haber una Planilla de Venta Activa")
          Call PINTATEXT(TXTNUMPLAVENTA)
          GoTo 99
    End If
      
    If CantRegistros&("P_VTA_BOLSAS", "NROPVTA = " & XVALO(TXTNUMPLAVENTA)) < 1 Then
        Call COMUNI("Debe Seleccionar Una Planilla de Venta Que esté Correctamente Guardada en al Base de Datos")
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

    Call AGREGARRUTAIMAG_PVTA(NUMPVTA&, APRO%)
    If APRO% < 0 Then GoTo 99
    Call GUARDARARCHIVOSADJUNTOS_PVTA(NUMPVTA&)

   '
   
   
99 Command47.Enabled = True
End Sub


Sub AGREGARRUTAIMAG_PVTA(Optional FTEC&, Optional OK%)

'      CD1.FILTER = "(*.*)|*.*|"
'      CD1.ShowOpen
'      Ruta$ = CD1.FileName
'      If TRIM(Ruta$) = "" Then Exit Sub
'      '
'20    DESCRI$ = TRIM$(InputBox$("Ingrese Texto ", "Guardar Como", CODI$))
'      If DESCRI$ = "" Then Exit Sub
'      DESCRI$ = UCase$(DESCRI$)
'      If Len(DESCRI$) > 24 Then
'        Call COMUNI("Longitud de Descripción \Supera La Cantidad Máxima de Caracteres (24)")
'        GoTo 20
'      End If
'      On Error Resume Next
'      CARPETA_ORIGEN$ = Ruta$
'      POS1% = InStrRev(Ruta$, "\")
'      POS2% = InStrRev(Ruta$, ".")
'      CARPETA_DESTINO$ = CONTROL("CARPIMAG", "UBICIMAG") & "\" & TRIM$(DESCRI$) & Mid$(Ruta$, POS2%)
'      FileCopy CARPETA_ORIGEN$, CARPETA_DESTINO$
'      If Err > 0 Then
'       A = 1
'      End If
'      On Error GoTo 0
'      'AGREGAR INFORMACION
'      Call REPLA(TX$, DESCRI$, 1, 24)
'      Call REPLA(TX$, CARPETA_DESTINO$, 32, 512)
'      Combo2.AddItem TX$

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
      CARPETA_DESTINO$ = Control("IMABOLSA", "UBIMABO") & "\" & TRIM$(DESCRI$) & Mid$(Ruta$, POS2%)
    
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

Private Sub Command45_Click()
   Command45.Enabled = False

   If Combo2.ListIndex >= 0 Then
     Linea$ = Combo2.List(Combo2.ListIndex)
     Archi$ = TRIM$(Mid$(Linea$, 25))
     If (EXISTE%(Archi$)) <> 1 Then
       Call COMUNI("Archivo Inexistente")
       GoTo 99
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

Sub Command5_Click()
   X = LISTASELECGENERAL("ANCHOTOT", , , 1)
   Me.TXTANCHOTOTAL = TRIM$(RESP_ZELE_VECT(2))
   Call VALORESPORDEFECTO
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

 Sub Command6_Click()
    Command6.Enabled = False
    X = LISTASELECGENERAL("TABLASI_NO", , , 1)
    Me.TXTIMPRESION = TRIM$(RESP_ZELE_VECT(2))
    Call VALORESPORDEFECTO
99  Command6.Enabled = True
    '
End Sub

Private Sub Command7_Click()
   X = LISTASELECGENERAL("TABLASI_NO", , , 1)
   Me.TXTFUELLE = TRIM$(RESP_ZELE_VECT(2))
   Call VALORESPORDEFECTO

End Sub

Private Sub Command8_Click()
   X = LISTASELECGENERAL("TABLASI_NO", , , 1)
   Me.TXTMICROPERFORADO = TRIM$(RESP_ZELE_VECT(2))
   Call VALORESPORDEFECTO



End Sub


Private Sub Command9_Click()
    X = LISTASELECGENERAL("TABLASI_NO", , , 1)
    Me.TXTPLENO = TRIM$(RESP_ZELE_VECT(2))
    Call VALORESPORDEFECTO
End Sub

Private Sub Form_Load()
    '
    TXTFECHA = FECHATEX$(HOY(HO$))
    Call CREATABLA_P_VTA_BOLSAS
    Call mnuNewPV_Click 'CARGA COMO NUEVA PLANILLA DE VENTAS
    LBLNUEVAPVTA.Visible = True
    Me.Width = 10170
    MODIFIC% = 0
    '
End Sub

Private Sub mnuAbrirCotExistente_Click()
    NROCOTI& = NROCOTIBOLSAS(0, 0)
    Dim CTP As New Conteplast
    TXTNUMPLAVENTA = ""
    TXTNUMPLAVENTA = XVALO(RESP_ZELE_VECT(4))
    
    'VALIDAR SI ESTA APROBADA (DEBE TENER CODIGO)
    LBLNUEVAPVTA.Visible = False
    Call VALORESPORDEFECTO
    If TRIM$(Text5) <> "" Then
       Call BLOQUEAR_CONTROLES_ENCONTENEDOR(Me, 2, 1, 1, 1, 0, 1, 1, Me.Frame)
       Call BLOQUEAR_CONTROLES_ENCONTENEDOR(Me, 2, 1, 1, 1, 0, 1, 1, Me.Frame2(1))
       Call COMUNI("Ficha Técnica Aprobada, Deshabilitada Para Registrar Cambios")
       Command2.Enabled = True
       Command39.Enabled = False
       Command13.Enabled = False
    End If

    
    
    TXTNROCOTIZACION = NROCOTI&
    LBLNEWCOTIZACTP.Visible = False
    If CTP.CotiBolsaAprobada%(NROCOTI&) > 0 Then
       Call BLOQUEAR_CONTROLES_ENCONTENEDOR(Me, 2, 1, 1, 1, 0, 1, 1, Me.Frame)
       Call BLOQUEAR_CONTROLES_ENCONTENEDOR(Me, 2, 1, 1, 1, 0, 1, 1, Me.Frame2(1))
       Call BLOQUEAR_CONTROLES_ENCONTENEDOR(Me, 2, 1, 1, 1, 0, 1, 1, Me.Frame5)
       Call COMUNI("Cotización Aprobada La Misma, Deshabilitada Para Registrar Cambios")
       Command2.Enabled = True
    End If
    Me.Width = 14280
End Sub

Private Sub mnuAdjuntos_Click()
    If DERACCE%("Asignación de Carpeta de Imagenes", "CarpImag") < 2 Then Exit Sub
    SELCARPETA.CONTROL1 = "IMABOLSA"
    SELCARPETA.CONTROL2 = "UBIMABOL"
    SELCARPETA.Show 1

End Sub

Private Sub mnuAprobarFichBol_Click()
       
      NC = XVALO(TEXT1)
      If NC < 1 Then
         Call COMUNI("Falta Ingresar Número de Cliente")
         Exit Sub
      End If
      
      If TRIM$(Text5) <> "" Then
         Call COMUNI("La Ficha Técnica Activa Ya Está Aprobada")
         Screen.MousePointer = 1
         Exit Sub
      End If
      
      If XVALO(TXTNUMPLAVENTA) < 1 Then
         Call COMUNI("Falta Seleccionar Una Ficha Activa")
         Screen.MousePointer = 1
         Exit Sub
      End If
      
      If VALIDACIONGRABACION%(1) < 0 Then Exit Sub  '

      If FIRMAELECTRONICAOK% <= 0 Then Exit Sub
      Me.TXTAPROBO = USERNAM$(XVALO(FIRMAELEC.Label1))
      Me.TXTFECHAPROBO = FECHATEX$(HOY(HO$))
      

      CI1% = 0:
      
      NUEVOCODIGO$ = FORMATOCODIGOCONTEPLASTBOLSAS$(NC)
      '
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
      On Error GoTo ERROR_GUARDAR
      FLEXCONN.BeginTrans

      Call ALTACODIGO(VDEF$, OKX%)
      If OKX% < 1 Then GoTo 80
      If OKX% = 1 Then
          CODINUE$ = TRIM$(Mid$(VDEF$, 77, LONCODI%))
          Text5.TEXT = CODINUE$
          'GRABO POR QUE EN EL OBJETO NO ENTRAN LOS 64 CARACTERES Y LO GRABA TRUNCADO
          SQLX$ = "UPDATE MASTER SET  Timate =  5, LINDIVIS =  1 , STOMAX = 1, FAMILIA = 'BO'"
          SQLX$ = SQLX$ + " WHERE CODINT = " & CI_NEW%
          FLEXCONN.Execute (SQLX$)
       End If
       FTEC& = XVALO(TXTNUMPLAVENTA)
       CONDI$ = "NroPvta = " & FTEC&
       SQLX$ = "UPDATE P_VTA_BOLSAS SET CODIMASTER = '" & CODINUE$ & "' ,NUCLIEN = " & NC
       SQLX$ = SQLX$ + " ,APROBO = '" & TXTAPROBO & "' , FEAPROBO = '" & CVDAT(HOY(HO$)) & "' ,STATUS = 3 ,CODINT = " & CI_NEW%
       SQLX$ = SQLX$ + " WHERE " & CONDI$
       FLEXCONN.Execute (SQLX$)
       '
       Dim CTP As New Conteplast
       Call CTP.AgregarOperacionesSegúnCódigo(CI1%) 'le agrega la operacion unica "FABRICACIÓN DE BOLSA"

       TXTNUMPLAVENTA = ""
       TXTNUMPLAVENTA = FTEC&
       Unload FIRMAELEC
       Screen.MousePointer = 1
       
ERROR_GUARDAR:
    If Err <> 0 Then
80     A = Err.Description
       On Error Resume Next
       FLEXCONN.RollbackTrans
       On Error GoTo 0
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION")
       Screen.MousePointer = 1
       Exit Sub
    Else
      FLEXCONN.CommitTrans
      Call MnuGuardar_Click
      'REFRESCA LA PANTALLA
      NROFT& = XVALO(TXTNUMPLAVENTA)
      If NROFT& > 0 Then
          TXTNUMPLAVENTA = ""
          LBLNUEVAPVTA.Visible = False
          TXTNUMPLAVENTA = NROFT&
          Call VALORESPORDEFECTO
     End If


    End If
 

End Sub
Function FIRMAELECTRONICAOK%()
     FIRMAELEC.TEXT1 = ""
     FIRMAELEC.Text2 = ""
     FIRMAELEC.Label1 = ""
     FIRMAELECTRONICAOK% = -1
     FIRMAELEC.Show 1

     If FIRMAELEC.ESTADO > 0 Then
        FIRMAELECTRONICAOK% = 1
        Exit Function
     Else
        If NIVCONFI%(XVALO(FIRMAELEC.Label1)) < 3 Then
         Call COMUNI("Usuario Con Nivel de Clasificación Inválido Para Esta Operación")
         FIRMAELECTRONICAOK% = -1
         Exit Function
        End If
     End If

End Function
Function FORMATOCODIGOCONTEPLASTBOLSAS$(NC)
   FORMATOCODIGOCONTEPLASTBOLSAS$ = ""
   If NC < 1 Then Exit Function
   '123456-89-11
   NC1& = NC
   PARTENROCLIENTE$ = FORMATOCON0$(NC1&, 6)
   PARTEORDEN$ = FORMATOCON0$(NUEVOORDENCODIGOMASTERDELCLIENTEBOLSA%(NC1&), 2)
   
   PARCIALCODIGO$ = PARTENROCLIENTE$ & "-" & PARTEORDEN$ & "-"
   
   FORMATOCODIGOCONTEPLASTBOLSAS$ = "CTP-" & PARTENROCLIENTE$ & "-" & PARTEORDEN$
End Function
Private Sub mnuAPVE_Click()
   If VALIDASISEMODIFICO% < 0 Then Exit Sub
   NUS% = USNBR% Mod 100
   VENDE% = LOGVEND.VendeVsUsuario(NUS%)
   Call BLANFORMU
   NRO = Sel_P_Vta&()
   If NRO > 0 Then
     TXTNUMPLAVENTA = ""
     LBLNUEVAPVTA.Visible = False
     TXTNUMPLAVENTA = NRO
     Call VALORESPORDEFECTO
   End If
   MODIFIC% = 0
   If TRIM$(Text5) <> "" Then
       Dim CTP As New Conteplast
       'SE VALIDA SI LA FICHA TECNICA POSEE ALGUNA COTIZACION APROBADA, SI NO, PERMITE EDITARLA POR MAS QUE LA
       'FICHA TECNICA HAYA SIDO APROBADA, SOLICITUD DE FRANCO
       CONDI$ = "NROPVTA = " & NRO & " and STATUS = 3"
       NROCOTI& = XVALO(VALOBADA("COTIPVTA", "NROCOTIZACION", CONDI$, "NOMESS"))
       If CTP.CotiBolsaAprobada%(NROCOTI&) > 0 Or DERACCE("MODFTAPR", "MODIFICAR FICHA TÉCNICA DE BOLSAS APROBADA") < 2 Then
         Call BLOQUEAR_CONTROLES_ENCONTENEDOR(Me, 2, 1, 1, 1, 0, 1, 1, Me.Frame)
         Call BLOQUEAR_CONTROLES_ENCONTENEDOR(Me, 2, 1, 1, 1, 0, 1, 1, Me.Frame2(1))
         Call COMUNI("Ficha Técnica Aprobada, Deshabilitada Para Registrar Cambios, La Misma Tiene Cotizacion/es Aprobadas")
       End If
       
       Command2.Enabled = True
       Command39.Enabled = True
       Command13.Enabled = True
       Set CTP = Nothing

    End If
    Me.Width = 10170
   'SIEMPRE QUEDA HABILITADA LA POSIBILIDAD DE AGREGAR UN ARCHIVO

End Sub
Function Sel_P_Vta&()
    Erase RESP_ZELE_VECT
    Campos$ = "F.Téc.Bolsa/F11;Fecha/D8;Cuenta/A12;R.Social/A32;Cant./F8.1;Vend/F5;ID/A0;Código/A16"
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    FRMZELECHO.Caption = ""
    
    
    Dim obj As New RECORXET
    CAMPOSTABLA$ = "NROPVTA,FECHEMISI,NUCLIEN,RASOCLI,CANTIDAD,NUVEND,P_VTA_BOLSAS_ID,CODIMASTER"
    TABLA$ = "P_VTA_BOLSAS"
    CONDI$ = "NROPVTA > 0 "
    '
    CONDI$ = CONDI$ + " ORDER BY NROPVTA DESC"
    Set RS1 = obj.RS_GENERAL(TABLA$, CAMPOSTABLA$, CONDI$)
    Call Carga_MSF_Recordset(RS1, Campos$, FRMZELECHO.MSF2, 1)
    On Error Resume Next
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    On Error GoTo 0
        
    FRMZELECHO.Width = 14000
50  FRMZELECHO.Show 1
    If XVALO(RESP_ZELE_VECT(7)) > 0 Then ' 7 ES EL ID QUE SI SELECCIONO ES > 0 POR QUE ES LA CONDICION
       Sel_P_Vta& = XVALO(RESP_ZELE_VECT(1))
    End If
End Function

Private Sub mnuCostosCotizacion_Click()
    CostoCot.Show 1
End Sub


Private Sub MnuGuardar_Click()

'   If VALIDACIONGRABACION% < 0 Then Exit Sub
   If XVALO(TEXT1) < 1 Then
      Call COMUNI("Falta Número de Cliente")
      Call PINTATEXT(TXTNUMPLAVENTA)
      Exit Sub
   
   
   End If
   If XVALO(Me.TXTNUMPLAVENTA) < 1 Then
      Call COMUNI("Falta Número de Planilla Venta")
      Call PINTATEXT(TXTNUMPLAVENTA)
      Exit Sub
   End If

   If TRIM$(TXTRAZOCLI) = "" Then
      Call COMUNI("Falta Ingresar Razón Social")
      Call PINTATEXT(TXTRAZOCLI)
      Exit Sub
   End If
   
   ESNUEVO% = 0
   If LBLNUEVAPVTA.Visible = True Then
      ESNUEVO% = 1
   End If
   '
   SQLX$ = "SELECT * FROM P_VTA_BOLSAS "
   If ESNUEVO% < 1 Then
     SQLX$ = SQLX$ + " WHERE NROPVTA = " & XVALO(TXTNUMPLAVENTA)
   Else
     SQLX$ = SQLX$ + " WHERE NROPVTA < 1 "
   End If
   
   On Error GoTo ERROR_GUARDAR
   FLEXCONN.BeginTrans
   
   Dim RS As ADODB.Recordset
   Set RS = New ADODB.Recordset
   RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   With RS
     If ESNUEVO% > 0 Then
         NEWCOTIZ = NumNuevaPVta&
         .AddNew
         !NroPvta = NEWCOTIZ
         Else
           NEWCOTIZ = XVALO(TXTNUMPLAVENTA)
     End If
     
        NC = XVALO(TEXT1)
        !rasocli = Left$(TXTRAZOCLI, 128)
        !NUCLIEN = NC
        !OBSERVACIONES = Left$(TXTOBSERVA, 1024)
        !Cantidad = XVALO(Text2)
        !NUVEND = XVALO(Text4)
        !FECHEMISI = CVDAT(CVINT(TXTFECHA))
        !CAPKGS = XVALO(TXTCAPACIDAD)
        !IMPRESION = Left$(Me.TXTIMPRESION, 24)
        !CANTICARAS = XVALO(TXTCANTCARAS)
        !Status = 0
        !MAT_BOLSA = Left(Me.TXTMATERIALBOLSA, 128)
        !ALTOUTILCM = XVALO(TXTALTOUTIL)
        !ANCHOTOTCM = XVALO(Me.TXTANCHOTOTAL)
        !Producto = Left(Me.TXTPRODUCTO, 64)
        !CANTCOLORES = XVALO(Me.TXTCANTCOLORES)
        !Fuelle = Left(Me.TXTFUELLE, 24)
        !MICROPERFO = Left(TXTMICROPERFORADO, 24)
        !Pleno = Left(Me.TXTPLENO, 24)
        !MEDIFUELLE = XVALO(Me.TXTMEDFUELLE)
        !LINER = Left(Me.TXTLINER, 24)
        !CODIMASTER = TRIM$(Left(Me.Text5, 24))
        !CODINT = CODINT%(TRIM$(Text5))
        !UNIPALLET = XVALO(UNIPALLET)
        !SISDESCA = TRIM$(Left$(SISDESCA, 48))
        !SisCarga = TRIM$(Left$(SisCarga, 48))
        !TIPOHILO = TRIM$(Left$(TIPOHILO, 48))
        !UNIDPAQU = XVALO(UNIDPAQU)
        !DescarCOLOR = Left$(Text32, 512) 'COLORES DE SISTEMA DE DESCARGA
        !DescarCantCol = XVALO(Text8)     'COLORES DE SISTEMA DE DESCARGA
        !APROBO = TRIM$(TXTAPROBO)
        !FEAPROBO = CVDAT(CVINT(FEAPROBO))

        
    .Update
   End With
   
   TXTNUMPLAVENTA = NEWCOTIZ
    
ERROR_GUARDAR:
    If Err.Number <> 0 Then
       FLEXCONN.RollbackTrans
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description)
       On Error GoTo 0
       
    Else
      FLEXCONN.CommitTrans
      If NOPRESENTARMENSAJE% < 1 Then
         Call COMUNI("Transacción Realizada")
      End If
      Me.LBLNUEVAPVTA.Visible = False
      MODIFIC% = 0
      RS.Close
      Set RS = Nothing
    End If



End Sub


Private Sub mnuHabiltCot_Click()
   '******VALIDAR SI LA MISMA YA ESTA APROBADA
   CONDI$ = "NROCOTIZACION = " & XVALO(TXTNROCOTIZACION)
   If CantRegistros&("COTIZACTP", CONDI$, "NOMES") < 1 Then
      Call COMUNI("La Cotización Activa no ha Sido Grabada\Deberá Grabar la Misma para Luego Aprobarla")
      Exit Sub
   End If
   
   CONDI$ = CONDI$ + " AND STATUS < 3 "
   If CantRegistros&("COTIPVTA", CONDI$, "NOMES") > 0 Then
      Call COMUNI("Cotización Activa No Aprobada\Al Aprobar Se Generará la Habilitación Para Poder Generar el Pedido")
      Exit Sub
   End If
   PorComisReal# = XVALO(Text9)
   CI1% = CODINT%(Text5)
   NVEND% = XVALO(Text4)
   Codigo$ = Text5
'   CODBIG$ = COTICONT.TXTCODCONFECCION
   NROFICHATEC& = TXTNUMPLAVENTA
   NC = XVALO(TEXT1)
   A = rasocli$(1) 'para refrescar el recordset
   
   If TRIM$(rasocli$(NC)) = "" Then
      Call COMUNI("Cliente No Válido")
      Exit Sub
   End If

   If CI1% < 1 Then
      Call COMUNI("Código No Válido")
      Exit Sub
   End If
   If NVEND% < 1 Then
      Call COMUNI("Vendedor No Válido")
      Exit Sub
   End If
   If NROFICHATEC& < 1 Then
      Call COMUNI("Número Ficha Técnica No Válido")
      Exit Sub
   End If
   If PorComisReal# < 1 Then
      Call COMUNI("Falta Porcentaje de Comisiòn")
   End If
   
   
   FIRMAELEC.TEXT1 = ""
   FIRMAELEC.Text2 = ""
   FIRMAELEC.Label1 = ""
   If XVALO(TXTNROCOTIZACION) < 1 Then
      Call COMUNI("Debe Haber una Cotización Activa")
      Exit Sub
   End If

   FIRMAELEC.Show 1

   If FIRMAELEC.ESTADO < 1 Then
      Exit Sub
   End If
   
   'TRANSACCIÓN GRABACION
   On Error GoTo ERROR_GUARDAR
   FLEXCONN.BeginTrans
   Call FRMCOTI2.GRABARENPRESUVEN(NC, XVALO(TXTNROCOTIZACION), 1, CI1%, NVEND%, Codigo$, "", NROFICHATEC&, 2, XVALO(TXTPVTAREAL), Text7, OKX%, 1, PorComisReal#)

   If OKX% < 0 Then GoTo 70
    
ERROR_GUARDAR:
   If Err <> 0 Then
70     FLEXCONN.RollbackTrans

       On Error GoTo 0
       MsgBox Err.Description
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION")
       Exit Sub
   Else
      FLEXCONN.CommitTrans
      
      Call COMUNI("Grabación Realizado Exitosamente")
   End If

   
End Sub
Private Sub mnuNewPV_Click()
    Call VALIDASISEMODIFICO
    Call BLANFORMU
    Call BLOQUEAR_CONTROLES_ENCONTENEDOR(Me, 1, 1, 1, 1, 0, 1, 1, Me.Frame)
    Call BLOQUEAR_CONTROLES_ENCONTENEDOR(Me, 1, 1, 1, 1, 0, 1, 1, Me.Frame2(1))

    LBLNUEVAPVTA.Visible = True
    TXTNUMPLAVENTA = SAFETEXT$(NumNuevaPVta&)
    TXTFECHA = FECHATEX$(HOY(HO$))
    MODIFIC% = 0
    NUS% = USNBR% Mod 100
    Text4 = LOGVEND.VendeVsUsuario(NUS%)

End Sub

Private Sub mnuPrintPV_Click()
   Call PrintPlanillaVenta
End Sub

Private Sub mnuSolCotPend_Click()

   
    
    NroPvta& = NROSOLICITUDPVTA(1, NC, FT&)
    If NroPvta& < 1 Then GoTo 99
    
    If DERACCE("COTIPLVE", "COTIZAR PLANILLA DE VENTAS ACTIVA") < 2 Then GoTo 99

    FT& = XVALO(RESP_ZELE_VECT(1))
    NC = XVALO(RESP_ZELE_VECT(4))
    TXTNUMPLAVENTA = FT&

    If FT& < 1 Then
       Call COMUNI("Debe Haber Una Ficha Técnica Activa")
       GoTo 99
    End If
    
    If NC < 1 Then
       Call COMUNI("Falta Seleccionar Cliente")
       GoTo 99
    End If
    '
    If TRIM$(Text5) = "" Then
      Call COMUNI("Falta Código (Debe Aprobar la Ficha Técnica Para Generar el Mismo")
        GoTo 99
    End If

    '
    If CantRegistros&("P_VTA_BOLSAS", "NROPVTA = " & FT&) < 1 Then
        Call COMUNI("Debe Seleccionar Una Planilla de Venta Que esté Correctamente Guardada en al Base de Datos")
        GoTo 99
    End If
''''''''''''''''
    'VALIDAR SI ESTA APROBADA (DEBE TENER CODIGO)
    TXTNUMPLAVENTA = ""
    LBLNUEVAPVTA.Visible = False
    TXTNUMPLAVENTA = FT&
    Call VALORESPORDEFECTO
    If TRIM$(Text5) <> "" Then
       Call BLOQUEAR_CONTROLES_ENCONTENEDOR(Me, 2, 1, 1, 1, 0, 1, 1, Me.Frame)
       Call BLOQUEAR_CONTROLES_ENCONTENEDOR(Me, 2, 1, 1, 1, 0, 1, 1, Me.Frame2(1))
       Call COMUNI("Ficha Técnica Aprobada, Deshabilitada Para Registrar Cambios")
       Command2.Enabled = True
       Command39.Enabled = False
       Command13.Enabled = False
    Else
       Command39.Enabled = True
       Command13.Enabled = True
    End If


'''''''''''''''''''''


    Call BLANFORMUCOTI
'   Campos$ = "N° Planilla Vta./f16;N° Vend./F9;Ref-Código/A48;Cta./f6;Raz.Social/A32;ID Sol./f8"

    NVEN% = XVALO(RESP_ZELE_VECT(2))
    REFECODI$ = RESP_ZELE_VECT(3)
    NC = XVALO(RESP_ZELE_VECT(4))
    RS$ = XVALO(RESP_ZELE_VECT(5))
    IDSOL& = XVALO(RESP_ZELE_VECT(6))
    LBLIDSOLICITUDCOTIZACION = IDSOL&
    TXTNROCOTIZACION = NumNuevaCotiza 'NumNuevaCotizaPVTA
    Me.TXTFECHACOTI = FECHATEX$(HOY(HO$))
    LBLNEWCOTIZACTP.Visible = True
    Me.Width = 14280
    MODIFIC% = 0
    On Error Resume Next
    Call ACTUALIZAVALORES
    TXTCOSTOPAPEL.SetFocus
    Me.TXTGRAMAJE.SetFocus
    On Error GoTo 0

99
End Sub

Private Sub Text1_Change()
   NCLIE = XVALO(TEXT1.TEXT)
   AA = rasocli$(1)
   If NCLIE > 0 And NCLIE < 99000 Then
      Me.TXTRAZOCLI.BackColor = &HC0C0C0
      Me.TXTRAZOCLI.TEXT = rasocli$(NCLIE)
   Else
       If NCLIE = 99000 And Me.TXTRAZOCLI.BackColor = &HC0C0C0 Then
         Me.TXTRAZOCLI = ""
       Else
         Me.TXTRAZOCLI.BackColor = &H80FFFF
       End If
   End If
   MODIFIC% = 1
End Sub


Private Sub Text2_Change()
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

Private Sub Text32_Change()
   MODIFIC% = 1
   If TRIM$(Text32) = "" Then
     Text8.Locked = False
     Text8 = ""
     Text8.BackColor = &H80FFFF
   Else
     Text8.Locked = True
     Text8 = ""
     Text8.BackColor = &HC0C0C0
     Call TEXTOLOTES(Text32, ";")
     On Error Resume Next
     If UBound(CADENATEX$) > 0 Then
        Text8 = UBound(CADENATEX$)
     Else
        Text8 = ""
     End If
   End If
End Sub

Private Sub Text32_DblClick()
   Text32 = ""
   Text8 = ""
   Text8.BackColor = &H80FFFF
   Command52.Enabled = True

End Sub

Private Sub Text4_Change()
  MODIFIC% = 1
  Text3 = ECOARCH$("VENDEDOR", Chr$(XVALO(Text4.TEXT)))
  MODIFIC% = 1
End Sub


Private Sub Text9_LostFocus()
  Text9 = TRIM$(FORMATNUM$(XVALO(Text9), "F5.1"))
End Sub


Private Sub TXTALTOUTIL_LostFocus()
    Call VALORESPORDEFECTO
End Sub


Private Sub TXTANCHOPAPEL_LostFocus()
If XVALO(TXTANCHOPAPEL) = 0 Then TXTANCHOPAPEL = "0"
End Sub


Private Sub TXTCAPACIDAD_Change()
   MODIFIC% = 1
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
    TXTCAPACIDAD = TRIM$(FORMATNUM$(XVALO(TXTCAPACIDAD), "f9.0"))
End Sub

Sub VALORESPORDEFECTO()
    TXTLABEL(14) = "Pleno"
    Text6 = "Máximo de Colores"
    If TRIM(UCase$(TXTIMPRESION)) = "SI" Then
       TXTLABEL(14) = "Pleno"
    ElseIf TRIM(TXTIMPRESION) <> "" Then
       TXTLABEL(14) = ""
       Text6 = ""
       TXTMAXIMOCOLORES = ""
       TXTCANTCOLORES = ""
    End If
    '
    If TRIM(UCase$(TXTIMPRESION)) = "SI" Then
       TXTLABEL(16) = "Cant. Caras"
    ElseIf TRIM(TXTIMPRESION) <> "" Then
       TXTLABEL(16) = ""
    End If
    '
    If TRIM(UCase$(TXTIMPRESION)) = "SI" Then
       TXTLABEL(5) = "Cant. Colores"
    ElseIf TRIM(TXTIMPRESION) <> "" Then
       TXTLABEL(5) = ""
    End If


    If TRIM(UCase$(TXTFUELLE)) = "SI" Then
       TXTLABEL(17) = "Medida Fuelle(x1)Cm."
    ElseIf TRIM(TXTFUELLE) <> "" Then
       TXTLABEL(17) = ""
    End If
    '
    'INFROMACION SEGUN ARCHIVO ENVIADO POR RAUL X:\CONTEPLAST\FLEXOFT\EXCEL\AA PLANILLA PARA VENTAS FINAL.xlsx
    If TRIM(UCase$(TXTMATERIALBOLSA)) = "BOLSA POLIPAPEL" Then
       TXTLARGODECORTE = TRIM$(FORMATNUM$(XVALO(TXTALTOUTIL), "F12.0"))
    Else
       TXTLARGODECORTE = TRIM$(FORMATNUM$(XVALO(TXTALTOUTIL) + 4, "F12.0"))
    End If
    
    
    If TXTIMPRESION = "NO" Then
       RODILLOS = ""
    Else
        'TABLA DATASQL > RODICOLO >
        'CAMPO1= CODIGO
        'CAMPO2 =Rod -LPAP
        'CAMPO3 =EXTRA/F12
        'CAMPO4 =CantColor

        'BUSCA EL VALOR DEL LIMITE INFERIOR ES DECIR SI BUSCA EL 88 Y NO HAY VA A TRARE EL 86 EN CASO Q DE 86 SALTE A 90
        CONDI$ = "LISTA = 'RODICOLO' AND CONVERT(FLOAT, campo2)<= " & XVALO(TXTLARGODECORTE) & " ORDER BY CONVERT(FLOAT, campo2) DESC"
        VARLORLARGOCORTE2_INFERIOR = XVALO(VALOBADA("DATASQL", "TOP 1 CAMPO2", CONDI$, "NOMESS"))
        LC2 = VARLORLARGOCORTE2_INFERIOR
        If LC2 <= 0 Then 'SI NO ENCUENTRA <= BUSCA EL VALOR SIGUIENTE MAYOR
          CONDI$ = "LISTA = 'RODICOLO' AND CONVERT(FLOAT, campo2)>= " & XVALO(TXTLARGODECORTE) & " ORDER BY CONVERT(FLOAT, campo2) "
          VARLORLARGOCORTE2_SUPERIOR = XVALO(VALOBADA("DATASQL", "TOP 1 CAMPO2", CONDI$, "NOMESS"))
          LC2 = VARLORLARGOCORTE2_SUPERIOR
        End If
        TXTLARGODECORTE2 = TRIM$(FORMATNUM$(LC2, "F12.2"))
        '
        'CON LOS DATOS DE ARRIBA BUSCA EL VALOR PARA ASIGNAR A RODILLO
        If XVALO(TXTLARGODECORTE) = LC2 Then
           TXTRODILLO = TRIM$(FORMATNUM$(LC2, "F12.2"))
        Else
           'EN ESTE CASO LA CONDI ES POR IGUAL POR QUE ES UN DATO QUE VIENE EXISTENTE EN LA TABLA EN LAS CONSULTAS ANTERIORES
           CONDI$ = "LISTA = 'RODICOLO' AND CONVERT(FLOAT, campo2)= " & XVALO(LC2)
           TXTRODILLO = TRIM$(FORMATNUM$(LC2 + XVALO(VALOBADA("DATASQL", " CAMPO3", CONDI$, "NOMESS")), "F12.2"))
           
        End If
      End If
      
      CONDI$ = "LISTA = 'RODICOLO' AND CONVERT(FLOAT, campo2)= " & XVALO(TXTRODILLO)
      TXTMAXIMOCOLORES = Int(XVALO(VALOBADA("DATASQL", " CAMPO4", CONDI$, "NOMESS")))
      
      Call CALCULOPARACOTIZACION
  
    
End Sub


Private Sub TXTCOSTOCONFECCION_LostFocus()
If XVALO(TXTCOSTOCONFECCION) = 0 Then TXTCOSTOCONFECCION = "0"
End Sub


Private Sub TXTCOSTOIMPRESION_LostFocus()
If XVALO(TXTCOSTOIMPRESION) = 0 Then TXTCOSTOIMPRESION = "0"
End Sub


Private Sub TXTCOSTOPAPEL_LostFocus()
If XVALO(TXTCOSTOPAPEL) = 0 Then TXTCOSTOPAPEL = "0"
End Sub


Private Sub TXTGRAMAJE_Change()
  Call CALCULOPARACOTIZACION
End Sub

Private Sub TXTGRAMAJE_LostFocus()
   TXTGRAMAJE = FORMATNUM$((XVALO(TXTGRAMAJE)), "F12.0")
   If XVALO(TXTGRAMAJE) = 0 Then TXTGRAMAJE = "0"
End Sub


Private Sub TXTGRAMAJEPAPEL_LostFocus()
If XVALO(TXTGRAMAJEPAPEL) = 0 Then TXTGRAMAJEPAPEL = "0"
End Sub


Private Sub TXTLABEL_DblClick(Index As Integer)
    If Label1.Visible = False Then
        Label1.Visible = True
        TXTLARGODECORTE.Visible = True
        Label2.Visible = True
        TXTLARGODECORTE2.Visible = True
        Label3.Visible = True
        TXTRODILLO.Visible = True
     Else
        Label1.Visible = False
        TXTLARGODECORTE.Visible = False
        Label2.Visible = False
        TXTLARGODECORTE2.Visible = False
        Label3.Visible = False
        TXTRODILLO.Visible = False
      End If

End Sub


Private Sub TXTLARGOCALCULO_LostFocus()
   If XVALO(TXTLARGOCALCULO) = 0 Then TXTLARGOCALCULO = "0"
End Sub


Private Sub TXTLARGOCORTECOT_LostFocus()
   If XVALO(TXTLARGOCORTECOT) = 0 Then TXTLARGOCORTECOT = "0"
End Sub


Private Sub TXTLARGOCORTECOT2_LostFocus()
      If XVALO(TXTLARGOCORTECOT2) = 0 Then TXTLARGOCORTECOT2 = "0"
End Sub


Private Sub TXTLARGODECORTE_Change()
   TXTLARGOCORTECOT = Me.TXTLARGODECORTE
End Sub

Private Sub TXTLARGODECORTE2_Change()
   TXTLARGOCORTECOT2 = Me.TXTLARGODECORTE2
End Sub

Private Sub TXTLINER_Change()
'   If UCase$(TRIM$(Me.TXTLINER)) = "SI" Then
'     Frame3.Visible = True
'   Else
'     Frame3.Visible = False
'   End If
   
End Sub

Private Sub TXTNROCOTIZACION_Change()
    Call BLANFORMUCOTI
    MODIFIC% = 0
    If XVALO(TXTNROCOTIZACION) > 0 Then
      If CantRegistros&("COTIPVTA", "NROCOTIZACION = " & XVALO(TXTNROCOTIZACION)) < 1 Then
         ESNUEVO% = 1
         Exit Sub
      End If
   Else
      Exit Sub
   End If

   SQLX$ = "SELECT * FROM COTIPVTA  WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE NROCOTIZACION = " & XVALO(TXTNROCOTIZACION)
   Set RS = READSET(SQLX$)
   With RS
         TXTFECHACOTI = FECHATEX$(CVINT(!FECHEMISI))
         Text7 = SAFETEXT$(!OBSERVACIONES)
         TXTNROCOTIZACION = XVALO(!NROCOTIZACION)
         TXTPRECIOPP = XVALO(!PRECIO_PP)
         TXTPRECIOPAPEL = XVALO(!PRECIOPAPEL)
         TXTANCHOPAPEL = XVALO(!ANCHOPAPEL)
         TXTCOSTOIMPRESION = XVALO(!COSTOIMPRE)
         TXTCOSTOCONFECCION = XVALO(!COSTOCONFE)
         TXTRODILLO = XVALO(!RODILLOS)
         TXTPRECIOVENTA = XVALO(!PRECIOVENTA)
         TXTPVTAREAL = XVALO(!PREVTAREAL)
         TXTPESO = XVALO(!Peso)
         TXTGRAMAJEPAPEL = XVALO(!GRAMAJEPAPEL)
         TXTPESOPAPEL = XVALO(!PESOPAPEL)
         TXTLARGODECORTE = XVALO(!LARGOCORTE)
         TXTLARGOCORTECOT2 = XVALO(!LARGOCORTE2)
         TXTCOSTOPAPEL = XVALO(!COSTOPAPEL)
         TXTLARGOCALCULO = XVALO(!LARGOCAL)
         TXTGRAMAJE = TRIM$(XVALO(!GRAMAJE))
         Text9 = TRIM$(XVALO(!ComPorReal))
         
   End With
   RS.Close
   Set RS = Nothing
End Sub

Private Sub TXTNUMPLAVENTA_Change()
    Call BLANFORMU
    Me.Width = 10170 ' OCULTA LA COTIZACION.
    MODIFIC% = 0
    If XVALO(TXTNUMPLAVENTA) > 0 Then
      If CantRegistros&("P_VTA_BOLSAS", "NROPVTA = " & XVALO(TXTNUMPLAVENTA)) < 1 Then
         ESNUEVO% = 1
         Exit Sub
      End If
   Else
      Exit Sub
   End If

   SQLX$ = "SELECT * FROM P_VTA_BOLSAS  WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE NROPVTA = " & XVALO(TXTNUMPLAVENTA)
   Set RS = READSET(SQLX$)
   With RS
  
        TEXT1 = XVALO(!NUCLIEN)
        TXTRAZOCLI = SAFETEXT$(!rasocli)
        TXTOBSERVA = SAFETEXT$(!OBSERVACIONES)
        Text2 = XVALO(!Cantidad)
        Text4 = XVALO(!NUVEND)
        TXTFECHA = CVDAT(CVINT(!FECHEMISI))
        TXTCAPACIDAD = FORMATNUM$(XVALO(!CAPKGS), "F12.2")
        TXTIMPRESION = SAFETEXT$(!IMPRESION)
        TXTCANTCARAS = XVALO(!CANTICARAS)
        TXTMATERIALBOLSA = SAFETEXT$(!MAT_BOLSA)
        TXTALTOUTIL = XVALO(!ALTOUTILCM)
        TXTANCHOTOTAL = XVALO(!ANCHOTOTCM)
        TXTPRODUCTO = SAFETEXT$(!Producto)
        TXTCANTCOLORES = XVALO(!CANTCOLORES)
        TXTFUELLE = SAFETEXT$(!Fuelle)
        TXTMICROPERFORADO = SAFETEXT$(!MICROPERFO)
        TXTPLENO = SAFETEXT$(!Pleno)
        TXTMEDFUELLE = XVALO(!MEDIFUELLE)
        TXTLINER = SAFETEXT$(!LINER)
        Text5 = SAFETEXT$(!CODIMASTER)
        UNIPALLET = XVALO(!UNIPALLET)
        SISDESCA = TRIM$(!SISDESCA)
        SisCarga = TRIM$(!SisCarga)
        TIPOHILO = TRIM$(!TIPOHILO)
        UNIDPAQU = XVALO(!UNIDPAQU)
        Text32 = SAFETEXT$(!DescarCOLOR) 'COLORES DE SISTEMA DE DESCARGA
        Text8 = XVALO(!DescarCantCol)    'COLORES DE SISTEMA DE DESCARGA
        TXTAPROBO = SAFETEXT$(!APROBO)
        FEAPROBO = CVDAT(CVINT(!FEAPROBO))
        
   End With
   LBLNUEVAPVTA.Visible = False
   MODIFIC% = 0
   RS.Close
   Set RS = Nothing
   Call VALORESPORDEFECTO
   Call CARGARARCHIVOSENCOMBOPVTA(XVALO(TXTNUMPLAVENTA))
End Sub
Sub CARGARARCHIVOSENCOMBOPVTA(NUMPVTA&)
      Combo2.Clear
      SQLX$ = "SELECT * FROM DOCURELPVTA WITH(NOLOCK) WHERE NROPVTA  = " & NUMPVTA&
      Dim DOCTMP As ADODB.Recordset
      Set DOCTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
      With DOCTMP
        Do While Not .EOF
           DESCRI$ = SAFETEXT$(!Descripcion)
           RUTAX$ = SAFETEXT$(!Ruta$)
           Call REPLA(TX$, DESCRI$, 1, 24)
           Call REPLA(TX$, RUTAX$, 32, 512)
           Combo2.AddItem TX$
           .MoveNext
        Loop
      End With
End Sub

Private Sub TXTPESO_LostFocus()
If XVALO(TXTPESO) = 0 Then TXTPESO = "0"
End Sub


Private Sub TXTPESOPAPEL_LostFocus()
If XVALO(TXTPESOPAPEL) = 0 Then TXTPESOPAPEL = "0"
End Sub


Private Sub TXTPRECIOPAPEL_LostFocus()
   If XVALO(TXTPRECIOPAPEL) = 0 Then TXTPRECIOPAPEL = "0"
End Sub


Private Sub TXTPRECIOPP_LostFocus()
If XVALO(TXTPRECIOPP) = 0 Then TXTPRECIOPP = "0"
End Sub


Private Sub TXTPRECIOVENTA_LostFocus()
If XVALO(TXTPRECIOVENTA) = 0 Then TXTPRECIOVENTA = "0"
End Sub


Private Sub TXTPVTAREAL_LostFocus()
   TXTPVTAREAL = TRIM$(FORMATNUM$(XVALO(TXTPVTAREAL), "F12.3"))

End Sub


Private Sub TXTRAZOCLI_Change()
MODIFIC% = 1
End Sub


Private Sub TXTRAZOCLI_DblClick()
    If TXTRAZOCLI.BackColor = &HE0E0E0 Then
       TXTRAZOCLI.BackColor = &H80FFFF
       TXTRAZOCLI.Locked = False
       'Call PINTATEXT(TXTRAZOCLI)
       Command1.Enabled = False
       If XVALO(TEXT1.TEXT) <> 99000 Then TEXT1.TEXT = ""
    Else
       TXTRAZOCLI.BackColor = &HE0E0E0
       TXTRAZOCLI.Locked = True
       Command1.Enabled = True
    End If
    MODIFIC% = 1
End Sub


Private Sub TXTRAZOCLI_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If
End Sub

Sub PrintPlanillaVenta()
    Call SETPRINPORT("RESTORE")
    PUERTOIMPRE$ = Control$(IDENTER$, "PRINTSER")
    If PUERTOIMPRE$ = "" Then PUERTOIMPRE$ = Control$("*", "PRINTSER")
    If PUERTOIMPRE$ <> "" Then
       Call SETPRINPORT(PUERTOIMPRE$)
    End If
    '
    RUTAEXCEL$ = Control("COTICONT", "EXCEPVTA")
    If EXISTE(RUTAEXCEL$) < 1 Then
       Call COMUNI("Falta Configurar Ruta de Plantilla Excel de Formulario de Ficha Técnica")
       Exit Sub
    End If
    '

    With VENTABNEW.MSF
       .Rows = 30: .Cols = 3
       '.TextMatrix(1, 1) = TXTRAZOCLI 'nombre cliente
       .TextMatrix(1, 2) = TXTRAZOCLI 'nombre cliente
       '.TextMatrix(2, 1) = Text2
       .TextMatrix(2, 2) = Text2
       .TextMatrix(3, 2) = Text3
       .TextMatrix(3, 1) = Text3
       .TextMatrix(4, 1) = TXTLABEL(1)
       .TextMatrix(4, 2) = TXTMATERIALBOLSA
       .TextMatrix(5, 1) = TXTLABEL(2)
       .TextMatrix(5, 2) = TXTANCHOTOTAL
       .TextMatrix(6, 1) = TXTLABEL(40)
       .TextMatrix(6, 2) = TXTCAPACIDAD
       
       .TextMatrix(7, 1) = TXTLABEL(4)
       .TextMatrix(7, 2) = TXTIMPRESION
       .TextMatrix(8, 1) = TXTLABEL(5)
       .TextMatrix(8, 2) = TXTCANTCOLORES
       .TextMatrix(9, 1) = TXTLABEL(6)  'Fuelle
       .TextMatrix(9, 2) = TXTFUELLE  'Fuelle
       .TextMatrix(10, 1) = TXTLABEL(9) 'Micro Perforado
       .TextMatrix(10, 2) = TXTMICROPERFORADO  'Micro Perforado
       .TextMatrix(11, 1) = TXTLABEL(10) 'Producto
       .TextMatrix(11, 2) = TXTPRODUCTO 'Producto
       .TextMatrix(12, 1) = TXTLABEL(13) 'Alto Util Cm.
       .TextMatrix(12, 2) = TXTALTOUTIL  'Alto Util Cm.
       .TextMatrix(13, 1) = TXTLABEL(14)  'Pleno
       .TextMatrix(13, 2) = TXTPLENO  'Pleno
       .TextMatrix(14, 1) = TXTLABEL(16) 'Cant.Caras
       .TextMatrix(14, 2) = TXTCANTCARAS 'Cant.Caras
       .TextMatrix(15, 1) = TXTLABEL(17) 'Medida Fuelle(x1) cm.
       .TextMatrix(15, 2) = TXTMEDFUELLE 'Medida Fuelle(x1) cm.
       .TextMatrix(16, 1) = TXTLABEL(7) 'Liner
       .TextMatrix(16, 2) = TXTLINER 'Liner
       '.TextMatrix(17, 1) = TXTOBSERVA  'Observaciones
       .TextMatrix(17, 2) = TXTOBSERVA  'Observaciones
       .TextMatrix(18, 1) = TXTNUMPLAVENTA   'nroplanilla
       .TextMatrix(18, 2) = TXTNUMPLAVENTA   'nroplanilla
       .TextMatrix(19, 1) = TXTFECHA   'fecha
       .TextMatrix(19, 2) = TXTFECHA   'fecha
       
       .TextMatrix(20, 2) = "SI"
       If Combo2.ListCount < 1 Then
          .TextMatrix(20, 2) = "NO"
       End If
'
       .TextMatrix(21, 1) = TXTLABEL(12)   'medidas liner 'falta en el excel
       .TextMatrix(21, 2) = TRIM$(Text2)   '
       .TextMatrix(22, 1) = TXTLABEL(41)   'tipo de hilo 'falta en el excel
       .TextMatrix(22, 2) = TIPOHILO   '
       .TextMatrix(23, 1) = TXTLABEL(43)   'Unid. x Paquete
       .TextMatrix(23, 2) = TRIM$(UNIDPAQU)   '
       .TextMatrix(24, 1) = TXTLABEL(1)   'Sistema de Carga
       .TextMatrix(24, 2) = TRIM$(SisCarga)    '
       .TextMatrix(25, 1) = TXTLABEL(39)   'Sistema de Descarga
       .TextMatrix(25, 2) = TRIM$(SISDESCA)    '
       .TextMatrix(26, 1) = TXTLABEL(53)   'Colores
       .TextMatrix(26, 2) = TRIM$(Text32)     '
       .TextMatrix(27, 1) = TXTLABEL(42)   'Unid. x Pallet
       .TextMatrix(27, 2) = TRIM$(UNIPALLET)     '
       .TextMatrix(28, 2) = TRIM$(Text5.TEXT)

    End With
    'OJO SI SE LLAMA A PANTALLA A VENTABNEW CUANDO SE ACEPTA BORRA LAS FILAS QUE ESTAN VACIAS
'    VENTABNEW.Show 1
        
    Call Param_Print_Excel(RUTAEXCEL$, "SIMULADOR", VENTABNEW.MSF, PUERTOIMPRE$, "1;2", 1, 10, "", 1, 0)                      'imprime en excel la 1,2 Y 3 columna de mi msf en la columna 1 a partir de la fila 11 en el excel (este graba E imprime)
    Unload VENTABNEW

    Call SETPRINPORT("RESTORE")





End Sub

Private Sub TXTRODILLO_Change()
   TXTRODILLOCOT = TXTRODILLO
End Sub

Private Sub TXTRODILLOCOT_LostFocus()
If XVALO(TXTRODILLOCOT) = 0 Then TXTRODILLOCOT = "0"
End Sub


