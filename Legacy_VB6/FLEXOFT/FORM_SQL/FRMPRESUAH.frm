VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FRMPRESUAHP 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Presupuestos"
   ClientHeight    =   8685
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   17115
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8685
   ScaleWidth      =   17115
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text37 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   13080
      Locked          =   -1  'True
      TabIndex        =   82
      TabStop         =   0   'False
      Text            =   "Últimas 3 Ventas"
      Top             =   5520
      Width           =   3855
   End
   Begin VB.TextBox Text36 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   13080
      Locked          =   -1  'True
      TabIndex        =   81
      TabStop         =   0   'False
      Text            =   "Últimas 3 Cotizaciones"
      Top             =   2760
      Width           =   3855
   End
   Begin VB.CommandButton Command7 
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
      Height          =   660
      Left            =   8670
      TabIndex        =   80
      ToolTipText     =   "Seleccion de Cliente"
      Top             =   1690
      Width           =   255
   End
   Begin VB.TextBox Text34 
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
      Height          =   285
      Left            =   10200
      Locked          =   -1  'True
      TabIndex        =   75
      TabStop         =   0   'False
      Text            =   "Sub-Total:"
      Top             =   7440
      Width           =   975
   End
   Begin VB.TextBox Text33 
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
      Left            =   11280
      Locked          =   -1  'True
      TabIndex        =   74
      TabStop         =   0   'False
      Top             =   7440
      Width           =   1635
   End
   Begin VB.TextBox Text32 
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
      Height          =   285
      Left            =   10200
      Locked          =   -1  'True
      TabIndex        =   73
      TabStop         =   0   'False
      Text            =   "I.V.A.:"
      Top             =   7680
      Width           =   975
   End
   Begin VB.TextBox Text31 
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
      Left            =   11280
      Locked          =   -1  'True
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   7680
      Width           =   1635
   End
   Begin VB.TextBox Text30 
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
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   70
      TabStop         =   0   'False
      Text            =   "Validez Hasta"
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
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
      Height          =   330
      Left            =   2040
      TabIndex        =   69
      Top             =   7440
      Width           =   255
   End
   Begin VB.TextBox Text29 
      Alignment       =   1  'Right Justify
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
      Height          =   315
      Left            =   1440
      TabIndex        =   68
      Top             =   7440
      Width           =   585
   End
   Begin VB.TextBox Text28 
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
      Left            =   2295
      Locked          =   -1  'True
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   7440
      Width           =   2850
   End
   Begin VB.TextBox Text27 
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
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   66
      TabStop         =   0   'False
      Text            =   "Cond. Pago"
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
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
      Height          =   330
      Left            =   2040
      TabIndex        =   65
      Top             =   6960
      Width           =   255
   End
   Begin VB.TextBox Text26 
      Alignment       =   1  'Right Justify
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
      Height          =   315
      Left            =   1440
      TabIndex        =   64
      Top             =   6960
      Width           =   585
   End
   Begin VB.TextBox Text19 
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
      Left            =   2295
      Locked          =   -1  'True
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   6960
      Width           =   2850
   End
   Begin VB.TextBox Text25 
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
      Height          =   285
      Left            =   10080
      Locked          =   -1  'True
      TabIndex        =   62
      TabStop         =   0   'False
      Text            =   "Sub-Total:"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.TextBox Total 
      Alignment       =   1  'Right Justify
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
      Index           =   1
      Left            =   11280
      TabIndex        =   59
      TabStop         =   0   'False
      Text            =   " "
      Top             =   6960
      Width           =   1620
   End
   Begin VB.TextBox PORDESCU 
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
      Left            =   9840
      MaxLength       =   14
      TabIndex        =   58
      Top             =   7080
      Width           =   1065
   End
   Begin VB.TextBox Total 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
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
      Left            =   11280
      TabIndex        =   57
      Text            =   " "
      Top             =   6600
      Width           =   1620
   End
   Begin VB.TextBox Text24 
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
      Height          =   380
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   56
      Text            =   "[F2] Presenta Lista de Códigos -  [F4] Edición de Precio de Venta -  [F5] Graba e Imprime Presupuesto Activo."
      Top             =   8280
      Width           =   12735
   End
   Begin VB.TextBox TICAMBIS 
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
      Height          =   300
      Left            =   5640
      MaxLength       =   14
      TabIndex        =   48
      Top             =   6480
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.PictureBox Picture999 
      Height          =   330
      Index           =   0
      Left            =   5400
      ScaleHeight     =   270
      ScaleWidth      =   1110
      TabIndex        =   45
      Top             =   1440
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.TextBox Text21 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   44
      TabStop         =   0   'False
      Text            =   "MODIFICACIÓN DE PRESUPUESTO"
      Top             =   2450
      Visible         =   0   'False
      Width           =   10290
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
      Height          =   285
      Left            =   11280
      Locked          =   -1  'True
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   7935
      Width           =   1635
   End
   Begin VB.TextBox Text10 
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
      Height          =   285
      Left            =   10200
      Locked          =   -1  'True
      TabIndex        =   39
      TabStop         =   0   'False
      Text            =   "Total:"
      Top             =   7935
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Selecc.Código"
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
      Left            =   120
      TabIndex        =   37
      ToolTipText     =   "Seleccion de Artículo (F2)"
      Top             =   2475
      Width           =   2490
   End
   Begin VB.TextBox Text39 
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
      Left            =   105
      Locked          =   -1  'True
      TabIndex        =   36
      TabStop         =   0   'False
      Text            =   "L. Precios"
      Top             =   6540
      Width           =   1215
   End
   Begin VB.TextBox Text22 
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
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   6540
      Width           =   2850
   End
   Begin VB.TextBox Text23 
      Alignment       =   1  'Right Justify
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
      Height          =   315
      Left            =   1425
      TabIndex        =   34
      Top             =   6540
      Width           =   585
   End
   Begin VB.CommandButton Command11 
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
      Height          =   330
      Left            =   2025
      TabIndex        =   33
      Top             =   6540
      Width           =   255
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   120
      TabIndex        =   21
      Top             =   1440
      Width           =   12840
      Begin VB.TextBox Text15 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   380
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   55
         ToolTipText     =   "DOBLE CLICK PARA MODO EDICIÓN"
         Top             =   540
         Width           =   1350
      End
      Begin VB.TextBox Text16 
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
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   54
         TabStop         =   0   'False
         Text            =   "Costo"
         Top             =   270
         Width           =   1350
      End
      Begin VB.TextBox TXTIMPORTE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   380
         Left            =   11400
         Locked          =   -1  'True
         TabIndex        =   32
         ToolTipText     =   "('+')  PARA AGREGAR ('-') PARA BORRAR"
         Top             =   540
         Width           =   1305
      End
      Begin VB.TextBox Text6 
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
         Left            =   11400
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Text            =   "Importe"
         Top             =   270
         Width           =   1305
      End
      Begin VB.TextBox TXTCANTIDAD 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   380
         Left            =   6300
         TabIndex        =   29
         TabStop         =   0   'False
         Text            =   "   "
         ToolTipText     =   "('+')  PARA AGREGAR ('-') PARA BORRAR"
         Top             =   540
         Width           =   930
      End
      Begin VB.TextBox TXTPREUNIT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   380
         Left            =   10080
         Locked          =   -1  'True
         TabIndex        =   28
         ToolTipText     =   "DOBLE CLICK PARA MODO EDICIÓN"
         Top             =   540
         Width           =   1350
      End
      Begin VB.TextBox txtcodigo 
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
         Height          =   380
         Left            =   120
         TabIndex        =   27
         Top             =   540
         Width           =   2355
      End
      Begin VB.TextBox Text12 
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
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Text            =   "Código"
         Top             =   270
         Width           =   2355
      End
      Begin VB.TextBox Text11 
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
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Text            =   "Descripción"
         Top             =   270
         Width           =   3855
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
         Height          =   285
         Left            =   6300
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Text            =   "Cantidad"
         Top             =   270
         Width           =   930
      End
      Begin VB.TextBox Text9 
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
         Left            =   10080
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Text            =   "P.Venta"
         Top             =   270
         Width           =   1350
      End
      Begin VB.TextBox TXTDESCRIPCION 
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
         Height          =   380
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   540
         Width           =   3855
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Graba e Imprime"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15000
      TabIndex        =   17
      Top             =   8280
      Width           =   1995
   End
   Begin VB.TextBox TEXT13 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   5175
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   6930
      Width           =   4500
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Height          =   1575
      Left            =   120
      TabIndex        =   8
      Top             =   -120
      Width           =   12840
      Begin VB.CommandButton Command4 
         Caption         =   "Nuevo"
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
         Left            =   6975
         TabIndex        =   77
         ToolTipText     =   "Seleccion de Cliente"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox Text35 
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
         Height          =   285
         Left            =   2640
         TabIndex        =   76
         Top             =   1080
         Width           =   4230
      End
      Begin VB.TextBox Text20 
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
         Left            =   12300
         Locked          =   -1  'True
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   1100
         Width           =   435
      End
      Begin VB.TextBox Text17 
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
         Left            =   10500
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1100
         Width           =   1365
      End
      Begin VB.TextBox Text8 
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
         Left            =   8055
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   220
         Width           =   2055
      End
      Begin VB.TextBox Text5 
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
         Left            =   8055
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   505
         Width           =   4680
      End
      Begin VB.TextBox Text18 
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
         Left            =   8055
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   790
         Width           =   4680
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
         Height          =   285
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   795
         Width           =   4230
      End
      Begin VB.TextBox Text3 
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   505
         Width           =   4230
      End
      Begin VB.TextBox Text2 
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   220
         Width           =   4230
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00404040&
         Height          =   1120
         Left            =   120
         TabIndex        =   9
         Top             =   210
         Width           =   1400
         Begin VB.TextBox Text1 
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
            Left            =   80
            TabIndex        =   30
            Top             =   460
            Width           =   975
         End
         Begin VB.CommandButton Command19 
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
            Height          =   300
            Left            =   1100
            TabIndex        =   10
            ToolTipText     =   "Seleccion de Cliente"
            Top             =   480
            Width           =   250
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
            Height          =   240
            Left            =   120
            OleObjectBlob   =   "FRMPRESUAHP.frx":0000
            TabIndex        =   79
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label26 
            BackColor       =   &H00EEEEEE&
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
            Height          =   300
            Left            =   75
            TabIndex        =   71
            Top             =   765
            Width           =   1215
         End
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   240
         Left            =   1560
         OleObjectBlob   =   "FRMPRESUAHP.frx":006C
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   1560
         OleObjectBlob   =   "FRMPRESUAHP.frx":00D8
         TabIndex        =   12
         Top             =   525
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   1560
         OleObjectBlob   =   "FRMPRESUAHP.frx":014A
         TabIndex        =   13
         Top             =   815
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   240
         Left            =   6975
         OleObjectBlob   =   "FRMPRESUAHP.frx":01BC
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
         Height          =   255
         Left            =   6975
         OleObjectBlob   =   "FRMPRESUAHP.frx":022C
         TabIndex        =   15
         Top             =   525
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
         Height          =   255
         Left            =   6975
         OleObjectBlob   =   "FRMPRESUAHP.frx":029A
         TabIndex        =   16
         Top             =   810
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   240
         Left            =   10185
         OleObjectBlob   =   "FRMPRESUAHP.frx":0302
         TabIndex        =   18
         Top             =   240
         Width           =   1065
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
         Height          =   285
         Left            =   9240
         OleObjectBlob   =   "FRMPRESUAHP.frx":036E
         TabIndex        =   19
         Top             =   1120
         Width           =   1200
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   11340
         TabIndex        =   20
         Top             =   195
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   16777215
         Format          =   16515073
         CurrentDate     =   40810
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   285
         Left            =   11900
         OleObjectBlob   =   "FRMPRESUAHP.frx":03E6
         TabIndex        =   43
         Top             =   1120
         Width           =   360
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   1560
         OleObjectBlob   =   "FRMPRESUAHP.frx":044C
         TabIndex        =   78
         Top             =   1080
         Width           =   1095
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   120
      OleObjectBlob   =   "FRMPRESUAHP.frx":04BC
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   270
      Left            =   5160
      OleObjectBlob   =   "FRMPRESUAHP.frx":06F0
      TabIndex        =   41
      Top             =   7800
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   7080
      OleObjectBlob   =   "FRMPRESUAHP.frx":0768
      TabIndex        =   46
      Top             =   2040
      Visible         =   0   'False
      Width           =   1365
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   3660
      Left            =   120
      TabIndex        =   50
      Top             =   2770
      Width           =   12780
      _ExtentX        =   22543
      _ExtentY        =   6456
      _Version        =   393216
      Cols            =   6
      BackColor       =   16705494
      BackColorFixed  =   16501405
      BackColorSel    =   16777088
      BackColorBkg    =   14737632
      GridColor       =   4210752
      GridColorFixed  =   14737632
      MousePointer    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSF2 
      Height          =   2460
      Left            =   13080
      TabIndex        =   51
      Top             =   120
      Width           =   3900
      _ExtentX        =   6879
      _ExtentY        =   4339
      _Version        =   393216
      Cols            =   6
      BackColor       =   16705494
      BackColorFixed  =   16501405
      BackColorSel    =   16777088
      BackColorBkg    =   14737632
      GridColor       =   4210752
      GridColorFixed  =   14737632
      MousePointer    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSF3 
      Height          =   2460
      Left            =   13080
      TabIndex        =   52
      Top             =   3000
      Width           =   3900
      _ExtentX        =   6879
      _ExtentY        =   4339
      _Version        =   393216
      Cols            =   6
      BackColor       =   16705494
      BackColorFixed  =   16501405
      BackColorSel    =   16777088
      BackColorBkg    =   14737632
      GridColor       =   4210752
      GridColorFixed  =   14737632
      MousePointer    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSF4 
      Height          =   2460
      Left            =   13080
      TabIndex        =   53
      Top             =   5760
      Width           =   3900
      _ExtentX        =   6879
      _ExtentY        =   4339
      _Version        =   393216
      Cols            =   6
      BackColor       =   16705494
      BackColorFixed  =   16501405
      BackColorSel    =   16777088
      BackColorBkg    =   14737632
      GridColor       =   4210752
      GridColorFixed  =   14737632
      MousePointer    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   270
      Left            =   9720
      OleObjectBlob   =   "FRMPRESUAHP.frx":07C6
      TabIndex        =   61
      Top             =   6880
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "%:"
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
      Left            =   10995
      TabIndex        =   60
      Top             =   7035
      Width           =   225
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "T-C:"
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
      Left            =   5160
      TabIndex        =   49
      Top             =   6495
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Label MONEMIS 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pesos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   6720
      TabIndex        =   47
      ToolTipText     =   "Doble-Click para 'Dólares'"
      Top             =   6480
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   195
      Left            =   2040
      TabIndex        =   38
      Top             =   600
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Menu mnu 
      Caption         =   "Menú"
      Begin VB.Menu mnuabmcte 
         Caption         =   "ABM Clientes"
      End
      Begin VB.Menu MNUEDITPRESUP 
         Caption         =   "Editar Presupuesto"
      End
   End
   Begin VB.Menu pxcopia 
      Caption         =   "Pres.X.Copia"
      Begin VB.Menu mnupcopia 
         Caption         =   "Pres.x.Copia"
      End
   End
   Begin VB.Menu mnuvis 
      Caption         =   "Ver"
      Begin VB.Menu mnuvisualiza 
         Caption         =   "Visualizar Presupuesto"
      End
   End
   Begin VB.Menu MnuConfig 
      Caption         =   "Configuración"
      Begin VB.Menu MnuTabEntr 
         Caption         =   "Tabla de Vias de Entrega"
      End
      Begin VB.Menu MnuArt_Prov 
         Caption         =   "Asociar Art. a Proveedor"
      End
      Begin VB.Menu MnuProv_Art 
         Caption         =   "Asociar Prov. a Artículo"
      End
      Begin VB.Menu mnuDemora 
         Caption         =   "Tabla de Demoras"
      End
   End
End
Attribute VB_Name = "FRMPRESUAHP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'HAY QUE DESPLEGAR DEL BOTON DE SELECCION DE CODIGO QUE MUESTRE LOS ITEMS DE LA LISTA DE PRECIOS
'ACTIVA , DEBERIA MOSTRAR LOS ITEMS CON PRECIO.
'VER TAMBIEN EL CASO EN QUE SE UTILICE LISTA 0 PARA QUE APAREZCAN TODOS LOS ITEMS EN EL BUSCADOR SIN PRECIO
'CARGA DE DETALLE POR TEXTO LIBRE

Sub GENERA_TABLA_RELACION_ART_PROV()
   Call CREACAMPO("", "RELACODPROV", "Num_Prov", 4, 0, 1)
   Call CREACAMPO("", "RELACODPROV", "Cod_inte", 4, 0, 1)
   Call CREACAMPO("", "RELACODPROV", "Precio", 7, 0)
   Call CREACAMPO("", "RELACODPROV", "Moneda", 3, 2)
   Call CREACAMPO("", "RELACODPROV", "Fecha", 8, 0)
   Call CREACAMPO("", "RELACODPROV", "MarBorra", 3, 0)
   Call CREACAMPO("", "RELACODPROV", "Via", 3, 0)
End Sub

Sub PRINTPRES()

    Call PRIPRESU(1)

End Sub
Sub PRIPRESU(VIPREVIA%)
    '
    FORIPRE$ = "PRSP-AQM"
    '
    If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") < 1 Then
       TX$ = "FALTA FORMULARIO DE IMPRESION DE REMITOS"
       TX$ = TX$ + "EL PRESUPUESTO SE HA GUARDADO, NO SALDRÁ IMPRESO"
       MsgBox TX$
       Exit Sub
    End If
    '
    Call LIMPIATETAB
    '
    TIPODOC$ = "Presupuesto" ' guardar el documento immpreso
'    NUMERI$ = "00000000" + TRIM$(Str$(Text17)): NUMERI$ = Right(NUMERI$, 8)
'    NUMERI$ = "PR-" & NUMERI$ & "-" & TRIM$(Text20)
    NUMERI$ = "00000000" + TRIM$(Str$(XVALO(Text17))): NUMERI$ = Right(NUMERI$, 8)
    VERSIOX$ = "00" & SAFETEXT$(XVALO(Text20)): VERSIOX$ = Right(VERSIOX$, 2)
    NUMERI$ = "PR-" & NUMERI$ + "-" & VERSIOX$

    NUMEDOC$ = NUMERI$
    
    'If VIPREVIA% = 1 Then
    '   CANCELPRINT% = 1
    'End If
    '
    'DATOS DE ENCABEZADO
    NC& = XVALO(Text1)
    
    CODPARAM$(1) = "FECH-EMI": DOCPARAM$(1) = FECHATEX$(CVINT(DTPicker1.Value))   'FECHA
    CODPARAM$(2) = "RASO-CLI": DOCPARAM$(2) = TRIM$(Text2)
    CODPARAM$(3) = "DOMI-CLI": DOCPARAM$(3) = TRIM$(Text3)
    CODPARAM$(4) = "LOCA-CLI": DOCPARAM$(4) = TRIM$(Text4)
    CODPARAM$(5) = "TELE-CLI": DOCPARAM$(5) = TRIM$(Text8)
    CODPARAM$(6) = "REF-MAT3":  DOCPARAM$(6) = TRIM$(Text18) 'MAIL
    CODPARAM$(7) = "REF-MAT4":  DOCPARAM$(7) = TRIM$(Text5) 'CELLUALR
    CODPARAM$(8) = "IMP-NETO":  DOCPARAM$(8) = TRIM$(Text14) 'SUBTOTAL
    CODPARAM$(9) = "IMP-TOTA":  DOCPARAM$(9) = TRIM$(Text14.TEXT) 'NONTO TOTAL
    CODPARAM$(10) = "NUME-COM":  DOCPARAM$(10) = NUMERI$ ' Text17 + 'NUMERO DE DOCUMENTO
        CODIPROVINCIA% = XVALO(VALOBADA("DEUDOR12", "Prov_Cli", CONDI$, "NOMESS"))
    CODPARAM$(11) = "PCIA-CLI":  DOCPARAM$(11) = ECOPROVINCIA$(CODIPROVINCIA%) 'PROVINCIA
    CODPARAM$(12) = "OBSERVA": DOCPARAM$(12) = TEXT13
    '
    CAPARDOC% = 12
    '
    CACOLTAB1% = 1
    CODTABU1$(1) = "CANTIDAD"
    CODTABU1$(2) = "COD-MPRI"
    CODTABU1$(3) = "DES-MAT"
    CODTABU1$(4) = "UNIMED"
    CODTABU1$(5) = "PRE-NETO"
    CODTABU1$(6) = "NET-ITEM"
    CODTABU1$(7) = "REF-MAT5"
    CODTABU1$(8) = "REF-MAT6"
    CACOLTAB1% = 8
    '
    CAITAB1% = 0
    TIOPCI% = 0 ' BANDERA PARA IMPRIMIR TITULO DE OPCIONALES
    For U& = 1 To MSF1.Rows - 1
        CODE$ = TRIM$(MSF1.TextMatrix(U&, 1))
        CIIX% = CODINT%(CODE$)
        DESCRIX$ = MSF1.TextMatrix(U&, 2)
        CAN = XVALO(MSF1.TextMatrix(U&, 3))
        PREUNI# = XVALO(MSF1.TextMatrix(U&, 4))
        IMPORTX# = XVALO(MSF1.TextMatrix(U&, 5))
        OPCIONAL% = XVALO(MSF1.TextMatrix(U&, 6))
        PRESENTA$ = SAFETEXT$(VALOBADA("MASTER", "PiezasXCaja", "CODINT=" & CIIX%))
        If PRESENTA$ = "0" Then PRESENTA$ = ""
        
        ' IMPRIME TITULO DE PRODUCTOS OPCIONALES
        If OPCIONAL% = 1 And TIOPCI% = 0 Then
            CAITAB1% = CAITAB1% + 1
            If U& > 1 Then
                CAITAB1% = CAITAB1% + 1
                For UU% = 1 To CACOLTAB1%: TETABU1$(UU%, CAITAB1% - 1) = "": Next UU%
            End If
            TETABU1$(8, CAITAB1%) = "Productos opcionales:"
            CAITAB1% = CAITAB1% + 1
            For UU% = 1 To CACOLTAB1%: TETABU1$(UU%, CAITAB1%) = "": Next UU%
            TIOPCI% = 1
        End If
        
        ' IMPRIME DETALLE, DISTINGUIENDO SI EL PRODUCTO ES OPCIONAL O NO
        CAITAB1% = CAITAB1% + 1
        If OPCIONAL% = 1 Then
            COD$ = AJUSTI(CODEXT$(CIIX%), 16)
            DES$ = AJUSTI(DESCRIT0$(CIIX%), 22)
            UME$ = AJUSTI(UNIMED$(CIIX%), 3)
            PST$ = AJUSTI(PRESENTA$, 8)
            QTD$ = FORMATNUM$(CAN, "F8.1")
            PRE$ = FORMATNUM$(PREUNI#, "F8.2")
            TOT$ = FORMATNUM$(CAN * PREUNI#, "F9.2")
            TETABU1$(8, CAITAB1%) = COD$ & " " & QTD$ & " " & UME$ & " " & PST$ & " " & DES$ & " " & PRE$ & " " & TOT$
        Else
            TETABU1$(1, CAITAB1%) = TRIM$(FORMATNUM$(CAN, "F8.1"))
            TETABU1$(2, CAITAB1%) = CODEXT$(CIIX%)
            TETABU1$(3, CAITAB1%) = DESCRIT0$(CIIX%)
            TETABU1$(4, CAITAB1%) = TRIM$(UNIMED$(CIIX%))
            TETABU1$(5, CAITAB1%) = FORMATNUM$(PREUNI#, "F12.2")
            TETABU1$(6, CAITAB1%) = FORMATNUM$(IMPORTX#, "F12.2")
            TETABU1$(7, CAITAB1%) = PRESENTA$
            TETABU1$(8, CAITAB1%) = ""
        End If
        '
        ' AGREGA LOS CODIGOS QUE COMPONEN LA FORMULA DEL PRODUCTO
        If CantRegistros("FORMULAS", "CODICONJ=" & CIIX%, "NOMESS") > 0 Then
            SQLX$ = "SELECT * FROM FORMULAS WITH(NOLOCK) WHERE CODICONJ=" & CIIX%
            Set RST = READSET(SQLX$)
            With RST
                Do While Not .EOF
                    CI0% = XVALO(!CODIELEM)
                    QUNI# = XVALO(!USOBRUTO)
                    CAITAB1% = CAITAB1% + 1
                    TETABU1$(1, CAITAB1%) = ""
                    TETABU1$(2, CAITAB1%) = ""
                    TETABU1$(3, CAITAB1%) = ""
                    TETABU1$(4, CAITAB1%) = ""
                    TETABU1$(5, CAITAB1%) = ""
                    TETABU1$(6, CAITAB1%) = ""
                    TETABU1$(7, CAITAB1%) = ""
                    TETABU1$(8, CAITAB1%) = ""
                    
                    If OPCIONAL% = 1 Then
                        DESXX$ = "... " & AJUSTI(DESCRIT0$(CI0%), 25) & " " & FORMATNUM$(QUNI#, "F10.2")
                        TETABU1$(8, CAITAB1%) = DESXX$
                    Else
                        DESXX$ = "... " & AJUSTI(DESCRIT0$(CI0%), 25) & " " & FORMATNUM$(QUNI#, "F10.2")
                        TETABU1$(3, CAITAB1%) = DESXX$
                    End If
                    .MoveNext
                Loop
            End With
        End If
        '
        Call CARLINSEP         'Carga linea de separacion entre codigo y codigo
        '
    Next U&
    '
    ' VERIFICA SI EL CLIENTE TIENE MAIL. EN CASO AFIRMATIVO CONSULTA AL USUARIO SI QUIERE ENVIARLO POR ESTA VIA
    SENDBYMAIL$ = ""
    EMAIL$ = TRIM$(SAFETEXT$(VALOBADA("DEUDOR12", "Mail_Cli", "NUME_CLI=" & NC&)))
    If EMAIL$ <> "" Then
        RSP% = MsgBox("¿Desea enviar el presupuesto por mail?", vbQuestion + vbYesNo, "Envio por mail")
        If RSP% = 6 Then SENDBYMAIL$ = EMAIL$
    End If
    
    Call PRINTDOC("@" + FORIPRE$)

    SENDBYMAIL$ = ""
    
    'BUSCA EL DOCUMENTO EN LA TABLA SQL CON LA NUEVA FUNCION
    DOCUNU& = DOCID&("PRESUPUESTO", NUMEDOC$, "")
5   Call MUESTRADOC(DOCUNU&)
    CANCELPRINT% = 0

99  Exit Sub
    '
End Sub


Private Sub Check1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        KeyAscii = 0
        Call LIMPIAR_TEXT_DETALLE
    End If
    
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        '
        'VALIDO EL CODIGO
        CI1% = CODINT%(txtcodigo)
        If CI1% < 1 Then
           MsgBox "CÓDIGO INEXISTENTE O NO VÁLIDO"
           Exit Sub
        End If

        'VALIDO QUE TENGA INGRESADA CANTIDAD
        If XVALO(TXTCANTIDAD.TEXT) < 0.05 Then
           MsgBox "FALTA CANTIDAD"
           TXTCANTIDAD.SetFocus
           Exit Sub
        End If
        'VALIDO QUE TENGA PRECIO
        If XVALO(TXTPREUNIT.TEXT) < 0.05 Then
           KeyAscii = 0
           MsgBox "FALTA PRECIO UNITARIO"
           TXTPREUNIT.SetFocus
           Exit Sub
        End If
        '
'        If VALIDA_COD_GRID(txtcodigo, FRMPRESU.MSF1, 1) > 0 Then
'         MsgBox "MODIFIQUE DATOS SOBRE EL MISMO"
'         GoTo 50
'        End If
        
        TXTPREUNIT.BackColor = &HE0E0E0 'CAMBIA A GRIS
        TXTPREUNIT.Locked = True 'DESHABILITO
        cargar_datos
        txtcodigo.SetFocus
50      Call LIMPIAR_TEXT_DETALLE

    End If
End Sub


Sub ULTIMOS3PRESUPUESTOS(MSF As MSFlexGrid, CI1&)
     MSF.Rows = 1
     SQLX$ = "SELECT TOP 3 NroPres, RasoClie, PreUnit,N "
     SQLX$ = SQLX$ + " FROM PRESUVEN WITH(NOLOCK)"
     SQLX$ = SQLX$ + " WHERE CODIINT =  " & CI1&
     SQLX$ = SQLX$ + " ORDER BY NROPRES DESC "
     Set RS = READSET(SQLX$)
     With RS
       Do While Not .EOF
        j& = j& + 1
        MSF.Rows = j& + 1
        MSF.TextMatrix(j&, 1) = XVALO(!NroPres)
        MSF.TextMatrix(j&, 2) = RASOCLI$(XVALO(!NroClie))
        MSF.TextMatrix(j&, 3) = TRIM$(FORMATNUM$(XVALO(!PreUnit), "F12.2"))
        .MoveNext
       Loop
     End With
     
End Sub

Sub Command1_Click()
   Command1.Enabled = False
   Call SELECHO("MASTER")
   COD$ = TRIM$(VALACT1$("MASTER"))
   If COD$ <> "" Then
     txtcodigo = COD$
   End If
   Command1.Enabled = True

End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   Call CARGA_INDICE_LPRECIO
   If RESPUESTA_BUSQUEDA$ <> "" Then Text23 = RESPUESTA_BUSQUEDA$
   Command11.Enabled = True

End Sub

Private Sub Command19_Click()
    If XVALO(Text20) > 0 Then Exit Sub 'si es modificacion
    Command19.Enabled = False
    Call SELECHO("DEUDOR1")
    NCLIE% = XVALO(TRIM$(VALACT1$("DEUDOR1")))
    If NCLIE% > 0 Then
      Text1.TEXT = TRIM$(Str$(NCLIE%))
      'Call CARDACLI
      On Error Resume Next
      txtcodigo.SetFocus
      On Error GoTo 0
    End If

    Screen.MousePointer = 1
    Command19.Enabled = True
    
    
 
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   If DERACCE%("ALTACLI", "Alta de Nuevo Cliente") < 2 Then
      Exit Sub
   End If
   '
   Call APERBASDAT("CLIEN")

   Dim Tranclie As ADODB.Recordset
   '
   Screen.MousePointer = 11
   PRONUCLI& = XVALO(Control$("", "PRONUCLI"))
   'If PRONUCLI& > 9999 Then PRONUCLI& = 0
   NROPRO& = PRONUCLI&
   If NROPRO& < 0 Then NROPRO& = 0
3  NROPRO& = NROPRO& + 1
   '
   'Set Tranclie = Badaclie.OpenRecordset("SELECT Nume_Cli FROM Deudor12 WHERE Nume_Cli = " & PRONUCLI&")
   'sql
   Set Tranclie = FLEXCONN.Execute(FILTSQL$("SELECT Nume_Cli FROM Deudor12 WHERE Nume_Cli = " & NROPRO&))
   If Not (Tranclie.EOF And Tranclie.BOF) Then 'si la encuentra (si el record no esta vacío)
      GoTo 3
   End If
'   '
   NCLIE = NROPRO&
   VDEF$ = String$(128, 0)
   Call REPLA(VDEF$, MKI$(Int(NCLIE)), 13, 2)
   Call REPLA(VDEF$, Space$(62), 15, 62)
   Call REPLA(VDEF$, Chr$(1), 77, 1)
   Call REPLA(VDEF$, Space$(15), 78, 15)
   '
10 Screen.MousePointer = 1
   OBX$ = OBJPLA$("ALTACLIENACC", VDEF$)
   If OBX$ = "" Then GoTo 99
   VDEF$ = OBX$
   '
   NCLINUE% = CVI(Mid$(OBX$, 13, 2))
   If NCLINUE% < 1 Or NCLINUE% > 29999 Then
      Call MENSERR(24, "Número de Cuenta no Válido.")
      GoTo 10
   End If
   '
   COCLINU$ = TRIM$(Left$(OBX$, 12))
   If COCLINU$ = "" Then
     COCLINU$ = "CLI-" + TRIM$(Str$(NCLINUE%))
     Call REPLA(VDEF$, COCLINU$, 1, 12)
     'Call MENSERR(24, "Codigo de Cuenta no Válido")
     'GoTo 10
   End If
   '
  
   Set Tranclie = FLEXCONN.Execute(FILTSQL$("SELECT Codi_Cli FROM Deudor12 where Codi_cli='" & COCLINU$ & "'"))
   If Not (Tranclie.EOF And Tranclie.BOF) Then
      Call MENSERR(24, "Ingreso no Válido.\Código de Cuenta Pre-existente.")
      GoTo 10
   End If
   '
   Set Tranclie = FLEXCONN.Execute(FILTSQL$("SELECT Nume_Cli FROM Deudor12 where Nume_cli=" & NCLINUE))
   If Not (Tranclie.EOF And Tranclie.BOF) Then
      Call MENSERR(24, "Ingreso no Válido.\Número de Cuenta Pre-existente.")
      GoTo 10
   End If
   '
   DCLINUE$ = TRIM$(Mid$(OBX$, 15, 62))
   If Len(DCLINUE$) < 1 Then
      Call MENSERR(24, "Ingreso no Válido.\Falta Razón Social del Cliente")
      GoTo 10
   End If
   '
   PIVCLI% = Asc(Mid$(OBX$, 77, 1))
   If PIVCLI% < 0 Or PIVCLI% > 7 Then
      Call MENSERR(24, "Ingreso no Válido.\Posición IVA no Figura en Tablas.")
      GoTo 10
   End If
   '
   CUTT$ = TRIM$(Mid$(OBX$, 78, 15))
   If VALICUIT%(CUTT$) <> 1 Then
     If PIVCLI% <> 0 Then
       If PIVCLI% <> 5 Then
         Call MENSERR(24, "Ingreso no Válido.\Número de C.U.I.T. Incorrecto.")
         GoTo 10
       End If
     End If
   End If
   '
   If PIVCLI% = 0 And Len(CUTT$) > 12 Then
        Call MENSERR(24, "El Límite Máximo Para Número de Documento es de 12 Dígitos")
        GoTo 10
   End If
   '
   If CUTT$ <> "" Then
     'sql
     Set Tranclie = FLEXCONN.Execute(FILTSQL$("SELECT Nume_Cli FROM Deudor12 WHERE Cuit_Cli = '" & CUTT$ & "'"))
     If Not (Tranclie.EOF And Tranclie.BOF) Then
         If Tranclie!NUME_CLI > 0 Then
            NCLIX% = Tranclie!NUME_CLI
            If COMALTER%("Existe Registrada una Cuenta con el Mismo C.U.I.T.\  \'" + TRIM$(RASOCLI$(NCLIX%)) + "' (" + TRIM$(Str$(NCLIX%)) + ")\\Cancelar\Registrar Igual") <> 2 Then
                GoTo 10
            End If
         End If
     End If
   End If
   '
   Screen.MousePointer = 11
   '
   Dim Rdeudor12 As ADODB.Recordset
   Set Rdeudor12 = New ADODB.Recordset
   Rdeudor12.CursorType = adOpenKeyset
   Rdeudor12.LockType = adLockPessimistic
   Rdeudor12.Open FILTSQL$("Select * from Deudor12 where nume_cli=" & NCLINUE%), FLEXCONN, , , adCmdText
   With Rdeudor12
      .AddNew   ' YA CONVERTIDO
      On Error Resume Next
      !CodiEmpr = CodiEmp%
      On Error GoTo 0
      !NUME_CLI = TRIM$(Str$(NCLINUE%))
      !Raso_Cli = DCLINUE$
      !Codi_Cli = COCLINU$
      !Cuit_Cli = CUTT$
      !Piva_Cli = PIVCLI%
      !Stat_Cli = 1
      '
      On Error Resume Next
      .Update
      If Err Then
         Call MENSERR(24, "Alta no Procesada Correctamente")
      End If
      On Error GoTo 0
   End With
   '
   Rdeudor12.Close
   Set Rdeudor12 = Nothing
   UPDATCLI% = 1
   
   Call CARGALISEL("DEUDOR1", "Deudor12", "Nume_Cli/F6.0>;Raso_Cli/A48", "Stat_Cli >= 0 and Nume_Cli>0 ORDER BY Raso_Cli")
   '
   Text1.TEXT = TRIM$(Str$(NCLINUE%))
   
99 Screen.MousePointer = 1
   Command4.Enabled = True
End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
   Text1.Locked = True
   Command19.Enabled = False
   CUENTA& = XVALO(Text1)
   If CUENTA& < 1 Then
      MsgBox "FALTA NÚMERO DE CLIENTE"
      GoTo 99
   End If
   '

   'VALIDACION DE DETALLE
   If MSF1.Rows <= 1 Then
     TTXX$ = "FALTA INGRESAR DETALLE DE PRESUPUESTO"
     MsgBox TTXX$
     GoTo 99
   End If
   '
   Call Grabar_Presupuesto
   '
   Call PRINTPRES
'   Call LIMPIARPANTALLA(Text1)
   Text1 = ""
   Text17 = 1 + MAXNUMCOD&("PRESUVENTA", "NroPres")
   Text20 = "-00"


99 Command5.Enabled = True
   Text1.Locked = False 'TEXT DE CUENTA DE CLIENTE
   Command19.Enabled = True 'BOTON DE LISTA DE SELECCIN
   If Text21.Visible = True Then Text21.Visible = False 'CARTEL DE MODIFICACION DE PRESUPUESTO
   '
End Sub

Sub Grabar_Presupuesto()
    '
''''    If Text21.Visible = False Then
''''      NUMPRES& = 1 + MAXNUMCOD&("PRESUVENTA", "NroPres")
''''      Text17 = NUMPRES&
''''    Else
''''      NUMPRES& = XVALO(Text17)
''''    End If
''''
''''    NUMERI$ = "00000000" + TRIM$(Str$(NUMPRES&)): NUMERI$ = Right(NUMERI$, 8)
''''    VERSIOX$ = "00" & SAFETEXT$(XVALO(Text20)): VERSIOX$ = Right(VERSIOX$, 2)
''''    NUMERI$ = "PR-" & NUMERI$ + "-" & VERSIOX$
''''
''''    SQLX$ = "SELECT * FROM PRESUVEN "
''''    SQLX$ = SQLX$ + " WHERE NroPres < 1"
''''    '
''''    On Error GoTo ERROR_GUARDAR
''''    Set aux = New ADODB.Recordset
''''    aux.Open (SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
''''    FLEXCONN.BeginTrans
''''    With aux
''''       .AddNew
''''       On Error Resume Next
''''       !CodiEmpr = CodiEmp%
''''       On Error GoTo 0
''''       '
''''       !NroPres = NUMPRES&
''''       !NroPres_Version = NUMERI$
''''       !Version = 0
''''       !FechEmis = DTPicker1.Value
''''       !NroClie = XVALO(Text1)
''''       !RASOCLI = Left$(Text2, 64)
''''       !Observa = TEXT13
''''       !Status = 0
''''       .Update
''''    End With
'''' '
''''    SQLX$ = "SELECT * FROM DETAPRESUPUESTO "
''''    SQLX$ = SQLX$ + " WHERE NroPres < 1"
''''
''''    Set aux = New ADODB.Recordset
''''    aux.Open (SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
''''    With aux
''''       For I& = 1 To MSF1.Rows - 1
''''        CODE$ = TRIM$(MSF1.TextMatrix(I&, 1))
''''        CI1% = CODINT%(CODE$)
''''        DESCRIX$ = MSF1.TextMatrix(I&, 2)
''''        CANTI = XVALO(MSF1.TextMatrix(I&, 3))
''''        PREUNI# = XVALO(MSF1.TextMatrix(I&, 4))
''''        IMPORTX# = XVALO(MSF1.TextMatrix(I&, 5))
''''        '
''''        .AddNew
''''        On Error Resume Next
''''        !CodiEmpr = CodiEmp%
''''        On Error GoTo 0
''''        '
''''        !NroPres = NUMPRES&
''''        !NroPres_Version = NUMERI$
''''        !Version = 0
''''        !FechEmis = DTPicker1.Value
''''        !NroClie = XVALO(Text1)
''''        !RASOCLI = Left$(Text2, 64)
''''        !Status = 0
''''        !MonedEmi = 1
''''        !NuOrdIt = I&
''''        !CodiInt = CI1%
''''        !CodiExt = CODE$
''''        !Descrip = DESCRIX$
''''        !PreListItem = 0
''''        !PreUnitItem = PREUNI#
''''        !Cantit = CANTI
''''        !ImToNeItem = IMPORTX#
''''        !Status = 0
''''        !CondVta = 0
''''        .Update
''''
''''       Next I&
''''    End With
''''
''''ERROR_GUARDAR:
''''    If Err <> 0 Then
''''      FLEXCONN.RollbackTrans
''''      MsgBox "ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description
''''      Exit Sub
''''    Else
''''      FLEXCONN.CommitTrans
''''      aux.Close
''''      Set aux = Nothing
''''
''''    End If
''''
    
    
    
    
    
    
    
    
    
    
    
    
   OKX% = 0
   Call BLOQARCH("PRESUPU")
   '
   NPP& = NUPRESUP&
   If NPP& > XVALO(Text17.TEXT) Then
      Call CIERRARCH("PRESUPU")
      Call COMUNI("Número Ocupado")
      Text11.TEXT = TRIM$(Str$(NPP&))
      Exit Sub
   End If
   '
   '\\\OT-06-0168-WAR-09/05/06///
   'validar no repeticion de codigo y numero de rubro
   Call ABREPRESVEN
   '\\\OT-06-0168-WAR-FIN///
   '
   If VALICLIE% < 1 Then
      MERRO$ = "Número de Cliente no Válido"
      GoTo 98
   End If
   NCLIE% = XVALO(Text1.TEXT)
   '
'   If Picture1.Visible = False Then
'     CIIX% = CODINT%(TEXT13.TEXT)
'     If CIIX% < 1 Or CIIX% > NUMAS% Then
'       MERRO$ = "Código de Artículo no Válido"
'       GoTo 98
'     End If
'   End If
   '
   If XVALO(Total(0).TEXT) < 1 Then
      MERRO$ = "Falta(n) Precio(s) de Lista"
      GoTo 98
   End If
   '
   If XVALO(Text14) < 1 Then
      MERRO$ = "Importe no Válido"
      GoTo 98
   End If
   '
   HOII% = HOY(HOS$)
   FEMI% = CVINT(DTPicker1.Value)
   If FEMI% < 1 Or FEMI% > HOII% Then
      MERRO$ = "Fecha de Emisión no Válida"
      GoTo 98
   End If
   '
   '
'   Call REPLA(XX$, Chr$(MONEMI%), 68, 1)
'   Call REPLA(XX$, MKS$(XVALO(Total(0).TEXT)), 71, 4)
'   Call REPLA(XX$, TRIM$(PORDESCU.TEXT), 75, 14)
'      DESC$ = PORDESCU.TEXT
'      TTT1 = COEFDESCU(DESC$)
'   Call REPLA(XX$, MKS$(TTT1), 89, 4)
'   Call REPLA(XX$, MKS$(XVALO(Total(1).TEXT)), 93, 4)
'   Call REPLA(XX$, MKS$(XVALO(Total(3).TEXT)), 97, 4)
'   Call REPLA(XX$, MKS$(1), 101, 4)
'   Call REPLA(XX$, MKD$(CDbl(1 * XVALO(Total(3).TEXT))), 105, 8)
   '
   NUOIT% = 0
   Call ABREPRESVEN
   'jandy sql
   STRSQL = "Select * from Presuven"
   STRSQL = STRSQL + " where NroPres=" & XVALO(Text11.TEXT)
   Dim PresuVen As ADODB.Recordset
   Set PresuVen = New ADODB.Recordset
   PresuVen.Open FILTSQL$(STRSQL), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   On Error Resume Next
   With PresuVen
   If Not (.EOF And .BOF) Then
      Call CIERRARCH("PRESUPU")
      Call COMUNI("Número Ocupado")
      Text11.TEXT = TRIM$(Str$(NPP&))
      Exit Sub
   End If
     
''''     CODICL$ = TRIM$(Label26) ' CARGA EL NUMERO DE PROSPECT
''''     Call REPLA(TTXYZ$, HOS$, 1, 8)
''''     Call REPLA(TTXYZ$, "Generación de Presupuesto:", 11, 56)
''''     Call REPLA(TTXYZ$, USERTER$, 69, 24)
''''     Call AGREGA_ANOTAPROSPECT(CODICL$, TTXYZ$) ' AGREGA EL ENCABEZADO EN EL ANOTADOR DEL PROSPECT
     
''''     'no lo encuentra entonces Inserta
''''     If Picture1.Visible = False Then
''''        '\\\OT-06-0168-WAR-09/05/06///
''''        CANTI0 = XVALO(Text14)
''''        If CANTI0 < 1 Then CANTI0 = 1
''''        '\\\OT-06-0168-WAR-FIN///
''''       .AddNew   ' YA CONVERTIDO
''''       On Error Resume Next
''''       !CodiEmpr = CodiEmp%
''''       On Error GoTo 0
''''        !NroPres = XVALO(Text11.TEXT)
''''        '\\\OT-06-0168-WAR-09/05/06///
''''         !NumeCRubro = 0
''''         !Cod_Rubro = ""
''''         !Cod_Anio = ""
''''        '\\\OT-06-0168-WAR-FIN///
''''        !FechEmis = CVDAT(FEMI%)
''''        !NroClie = NCLIE%
''''        !RasoClie = Left$(RASOCLI$(NCLIE%), 64)
''''        !Referen = Left$(Text10, 16)
''''        !AteSen = Left$(Text12, 48)
''''        !LisPre = XVALO(Text21(0))
''''        !ConPag = XVALO(Text21(1))
''''        !VENDED = XVALO(Text21(2))
''''        !MonedEmi = MONEMI%
''''        !NuOrdIt = NUOIT%
''''        !CodiInt = CIIX%
''''        !CodiExt = Left$(CODEXT$(CIIX%), 16)
''''        !Descrip = Left$(DESCRIT0$(CIIX%), 64)
''''        !Especif = Text28
''''        !PreList = XVALO(Total(0)) / CANTI0 'XVALO(Total(0))'\\\OT-06-0168-WAR-09/05/06///
''''        !TDESCUEN = Left$(PORDESCU, 32)
''''           DESC$ = PORDESCU
''''        !PorcDesc = COEFDESCU(DESC$)
''''        '
''''        If !PorcDesc = 0 Then
''''            !ImpoDesc = 0
''''        Else
''''            !ImpoDesc = XVALO(Total(1)) / CANTI0 'XVALO(Total(1))'\\\OT-06-0168-WAR-09/05/06///
''''        End If
''''        '
''''        !PreUnit = XVALO(Total(3)) / CANTI0  'XVALO(Total(3))'\\\OT-06-0168-WAR-09/05/06///
''''        !Cantit = CANTI0 '1'\\\OT-06-0168-WAR-09/05/06///
''''        !ImToNe = XVALO(Total(3)) '1 * XVALO(Total(3))'\\\OT-06-0168-WAR-09/05/06///
''''        !FechVent = CVDAT(0)
''''        !NroPedi = 0
''''        !NroFac_Cor = ""
''''        !NroFac_Lar = ""
''''        !Codi_Cli = CODICL$ ' GRABA CODIGO DE CUENTA DEL CLIENTE / PROSPECT
''''       .Update
''''        TTXYZ$ = Space$(76)
''''        Call REPLA(TTXYZ$, "- " + CODEXT$(CIIX%), 11, 18)
''''        Call REPLA(TTXYZ$, DESCRIT0$(CIIX%), 29, 38)
''''        Call REPLA(TTXYZ$, CSTRING$(MKS$(CANTI0), 3, 10, 1, 2), 57, 10)
''''        Call AGREGA_ANOTAPROSPECT(CODICL$, TTXYZ$) ' AGREGA EL DETALLE EN EL ANOTADOR DEL PROSPECT

''''      Else

        DESC$ = PORDESCU
        COEFIDESCU = COEFDESCU(DESC$)
       For U& = 1 To MSF1.Rows - 1
'         X$ = REGLEIDO$("!CARDETPR", U&)
         CIIX% = CODINT%(MSF1.TextMatrix(U&, 1))
         CAN = XVALO(MSF1.TextMatrix(U&, 3))
         INETO = XVALO(MSF1.TextMatrix(U&, 5)) * (100 - XVALO(COEFIDESCU)) / (100)

         If CAN > 0.05 Then
           .AddNew   ' YA CONVERTIDO
           On Error Resume Next
           !CodiEmpr = CodiEmp%
           On Error GoTo 0
             !NroPres = XVALO(Text17.TEXT)
             '\\\OT-06-0168-WAR-09/05/06///
             !NumeCRubro = 0
             !Cod_Rubro = ""
             !Cod_Anio = ""
             '\\\OT-06-0168-WAR-FIN///
             !FechEmis = CVDAT(FEMI%)
             !NroClie = NCLIE%
             !RasoClie = Left$(Text2, 64)
             !Referen = "" 'Left$(Text10, 16)
             !AteSen = "" 'Left$(Text12, 48)
             !LisPre = XVALO(Text23)
             !ConPag = XVALO(Text26)
             !VENDED = 0
             !MonedEmi = MONEMI%
             !NuOrdIt = NUOIT%
               NUOIT% = NUOIT% + 1
             !CodiInt = CIIX%
             !CodiExt = Left$(CODEXT$(CIIX%), 16)
             !Descrip = Left$(DESCRIT0$(CIIX%), 64)
             !Especif = DESCRIT0$(CIIX%)
             !PreList = INETO
             !TDESCUEN = Left$(PORDESCU, 32)
                DESC$ = PORDESCU
                TTT1 = COEFDESCU(DESC$)
             !PorcDesc = TTT1
             !ImpoDesc = INETO * TTT1 / 100
             !PreUnit = INETO * (1 - TTT1 / 100)
             !Cantit = CAN
             !ImToNe = CAN * INETO * (1 - TTT1 / 100)
             !FechVent = CVDAT(0)
             !NroPedi = 0
             !NroFac_Cor = ""
             !NroFac_Lar = ""
             !Codi_Cli = CODICL$
             !Vers_Presu = XVALO(Text20)
             .Update
            TTXYZ$ = Space$(76)
            Call REPLA(TTXYZ$, "- " + CODEXT$(CIIX%), 11, 18)
            Call REPLA(TTXYZ$, DESCRIT0$(CIIX%), 29, 38)
            Call REPLA(TTXYZ$, CSTRING$(MKS$(CAN), 3, 10, 1, 2), 57, 10)
            Call AGREGA_ANOTAPROSPECT(CODICL$, TTXYZ$) ' AGREGA EL DETALLE EN EL ANOTADOR DEL PROSPECT
         End If
       Next U&
''''     End If
   End With
   PresuVen.Close
   Set PresuVen = Nothing
   '
   Call OPLIPRE07.ACTINDIPRES(XVALO(Text11))
   Call CIERRARCH("PRESUPU")
   '
   '
   OKX% = 1
   Exit Sub
   '
98 Call CIERRARCH("PRESUPU")
   Call CIERRARCH("INDIPRES")
   Call MENSERR(24, MERRO$)
   '
End Sub


Function VALICLIE%()
   VCC% = 0
   If XVALO(Text1.TEXT) > 0 Then
     If XVALO(Text1.TEXT) <= 32767 Then
       NCLIE% = XVALO(Text1.TEXT)
       If REGICLI&(NCLIE%) > 0 Then
         VCC% = 1
       End If
     End If
   End If
   If VCC% = 0 Then
      CODIXX$ = TRIM$(Label26)
      CODICL$ = SAFETEXT$(VALOBADA("BSEGCOM", "Codi_Cli", "Codi_Cli ='" & CODIXX$ & "'"))
      If CODICL$ <> "" Then VCC% = 1
   End If
   '
99 VALICLIE% = VCC%
   '
End Function


Private Sub Command7_Click()
  
  PROVEEDOR , VIA, DEMORA

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode 'ESTO ES PARA QUE DESP
        Case 113 'pulso  f2
        Call Command1_Click
        Case 116 'pulso  f5
        Command5_Click

    End Select

End Sub

Private Sub Form_Load()

   'Call APLICAR_SKIN(Me, App.Path & "\SKINS\dogmax.skn")
   'Call iniciar_variables

    Call CENTRAFORM(Me)
'    Show
    Refresh
    '
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    '
15  EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
       EMPRE_SIN_USUARIO$ = REPLACAR$(EPAC$, USERNAME$, "")
       Caption = Caption + "  -  " + EMPRE_SIN_USUARIO$
    End If
    '
'    HOII% = HOY(HOS$)
'    Text9.TEXT = HOS$
    Text17.TEXT = TRIM$(Str$(NUPRESUP&))
    'Call NUCORUB
    '
    ARCHIPRE$ = "MASTER"
    '
    MONEMI% = 1  ' PESOS
    MONEMIS = DEMONECO$(1)
    MONEMIS.Visible = True
    TICAMBIS = FORMATNUM$(TICAMONE(1), "F9.4")
    MONEMIS.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MONEMI%), "F10.4")) + " - Doble-Click para Cambiar Moneda"
    CONVERPRECIO% = 1
    TICAMBIS.Visible = True
    Label11.Visible = True
    '
    For U& = 1 To 2
       X$ = REGLEIDO$("GRUPIVA", U&)
       PRIVAX = CVS(Mid$(X$, 13, 4))
       If U& = 1 Then
           PORIVA0 = TRIM$(FORMATNUM$(PRIVAX, "F8.2"))
         Else
           PORIVA1 = TRIM$(FORMATNUM$(PRIVAX, "F8.2"))
       End If
    Next U&
    Call CIERRARCH("GRUPIVA")
    '
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
    '
'
    DTPicker1.Value = Date
    
    Text17 = NUPRESUP&
    Text20 = "-00"
    '
'''    MSF1.Cols = 7
'''    MSF1.Rows = 1
'''    NUCOLU% = 1: TEPRUEBA = "12": MSF1.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU% - 1) = ALINEA1%: MSF1.TextMatrix(0, NUCOLU% - 1) = ""
'''    NUCOLU% = 2: TEPRUEBA = "123456789012345678901234": MSF1.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU% - 1) = ALINEA1%: MSF1.TextMatrix(0, NUCOLU% - 1) = "Codigo"
'''    NUCOLU% = 3: TEPRUEBA = "ABCDEFGHIJKLLLOPQRSTWXYZABCDEFGCHIHKLAAAAAAAAAAA": MSF1.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU% - 1) = 1: MSF1.TextMatrix(0, NUCOLU% - 1) = "Descripción"
'''    NUCOLU% = 4: TEPRUEBA = "1234567890123456789": MSF1.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU% - 1) = 6: MSF1.TextMatrix(0, NUCOLU% - 1) = "Cantidad"
'''    NUCOLU% = 5: TEPRUEBA = "1234567890123456789": MSF1.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU% - 1) = 6: MSF1.TextMatrix(0, NUCOLU% - 1) = "P.Unitario"
'''    NUCOLU% = 6: TEPRUEBA = "1234567890123456789": MSF1.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU% - 1) = 6: MSF1.TextMatrix(0, NUCOLU% - 1) = "Importe"
    Call CARGA_ENCABEZADO(FRMPRESUAHP.MSF1, "Código/A24;Descripción/A40;Cantidad/F9.1;Costo/F10.2;Pr.Venta/F10.2;Importe/F10.2", FRMPRESUAHP)
    Call CARGA_ENCABEZADO(FRMPRESUAHP.MSF3, "Presup./f8.0;Cliente/A20;P.Unit./F10.2", FRMPRESUAHP)
End Sub
Function NUPRESUP&()
   '
   NPP& = 1
   Call ABREPRESVEN
   '
   Dim MIRS As ADODB.Recordset
   STRSQL = ("SELECT max(NroPres) as maximo FROM PresuVen")
   Set MIRS = FLEXCONN.Execute(FILTSQL$(STRSQL))
   If Not IsNull(MIRS!Maximo) Then
      NPP& = 1 + MIRS!Maximo
   End If
   NUPRESUP& = NPP&
   MIRS.Close
   Set MIRS = Nothing
   '
End Function


Private Sub mnuedit_Click()

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If XVALO(Text20) > 0 Then
      'Call APLICAR_SKIN(Me, App.Path & "\SKINS\dogmax.skn")
      Text1.Locked = False 'TEXT DE CUENTA DE CLIENTE
      Command19.Enabled = True 'BOTON DE LISTA DE SELECCIN
      If Text21.Visible = True Then Text21.Visible = False 'CARTEL DE MODIFICACION DE PRESUPUESTO
      Text1 = ""
   Else
      Close
      End
   End If
End Sub

Sub CARGA_PRESUPUESTOS()
    Screen.MousePointer = 11
    RESPUESTA_BUSQUEDA$ = ""

    'VALIDO POR SI NO HAY NINGUN DATO CARGADO
    CONDI$ = "NroPres > 0 "
    CNRG& = CantRegistros&("PRESUVENTA", CONDI$)
    If CNRG& < 1 Then
      MsgBox "No Hay Presupuestos Cargados"
      Exit Sub
    End If

    '
    LISTASELEC.List1.Clear
    LISTASELEC.List2.Visible = False
    LISTASELEC.List1.Visible = True
    '
    'ASIGNO EL ANCHO QUE VA A TENER CADA CAMPO A ESTAS 2 VARIABLES PUBLICAS
    CODIGO_LONG% = 14
    DESCRIP_LONG% = 48
    
    SQLX$ = " SELECT * "
    SQLX$ = SQLX$ + " FROM PRESUVENTA "
    SQLX$ = SQLX$ + " ORDER BY NroPres "

    Set TABTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))

    With TABTMP
        Do While (.EOF = False)
          NPRX$ = SAFETEXT$(!NroPres_Version)
          NCLIE& = XVALO(!NroClie)
          RAZONCLI$ = SAFETEXT$(!RASOCLI)
          FEX = CVINT(!FechEmis)
          '
          YY2$ = " - " + AJUSTI$(RAZONCLI$, 20)
          YY2$ = YY2$ + " - " + AJUSTD$(FECHATEX$(FEX), 8)
          COCU$ = TRIM$(NPRX$)
          
          Call REPLA(TTXX$, COCU$, 1, CODIGO_LONG%)
          Call REPLA(TTXX$, YY2$, CODIGO_LONG% + 2, DESCRIP_LONG%)
          JJ& = JJ& + 1
          LISTASELEC.List1.AddItem TTXX$
          '
          If NPX& <> NPXA& Then
            JJ& = JJ& + 1
          End If
          NPXA& = NPX&
          '
        .MoveNext
        Loop
      
    End With
    '
    LISTASELEC.Caption = "INDICE DE PRESUPUESTOS"
    LISTASELEC.List1.ListIndex = LISTASELEC.List1.ListCount - 1
    Screen.MousePointer = 1
60  LISTASELEC.Show 1

End Sub


Private Sub mnuabmcte_Click()
    ABMCLIENTES.Show 1
End Sub

Private Sub MnuArt_Prov_Click()
      '
         Call GENERA_TABLA_RELACION_ART_PROV
5        Call SELECHO("MASTER") ' DESPLIEGA LA LISTADO DE ARTICULOS
         '
         COD$ = TRIM$(VALACT1$("MASTER"))
         If COD$ = "" Then Exit Sub
         '
         CI% = CODINT%(COD$)
         ' BUSCO EN LA TABLA PROVEDCOD TODOS LOS REGISTROS QUE COINCIDAN CON EL CI%
         SQLX$ = "SELECT * FROM RELACODPROV WITH (NOLOCK)"
         SQLX$ = SQLX$ + "WHERE COD_INTE = " & CI% & " AND MARBORRA = 0 "
         Dim RS As ADODB.Recordset
         Set RS = FLEXCONN.Execute(FILTSQL$(SQLX$))
         '
         Call SETULTREG("!PROVIADE", 0)
         JJ& = 0
         With RS
            Do While Not .EOF ' RECORRO LA TABLA Y GRABO EN EL ARCHIVO PARA MOSTRAR LA VENTANA
                X$ = REGBLAN$("PROVIADE")
                Call REPLA(X$, MKI(XVALO(!NUM_PROV)), 1, 2)
                Call REPLA(X$, MKI(XVALO(!VIA)), 3, 2)
                Call REPLA(X$, MKD(XVALO(!PRECIO)), 5, 8)
                Call REPLA(X$, MKI(XVALO(!MONEDA)), 13, 2)
                Call REPLA(X$, MKI(CVINT(!FECHA)), 15, 2)
                JJ& = JJ& + 1
                Call GRAREG("!PROVIADE", X$, JJ&)
            .MoveNext
            Loop
         End With
         RS.Close
         Set RS = Nothing
         
10       XYX% = VENTABU%("PROVIADE") ' MUESTRO LA VENTANA
         If XYX% < 0 Then GoTo 5     ' SI CIERRA VUELVE AL SELECHO
         '
         Screen.MousePointer = 11
         For U& = 1 To ULTREG&("!PROVIADE")
            X$ = REGLEIDO$("!PROVIADE", U&)
            COCO% = Asc(Left$(X$, 1))
            If COCO% < 1 Then
               Call MENSERR(24, "Código debe ser Numérico.")
               GoTo 10
            End If
            '
            DECO$ = TRIM$(Mid$(X$, 2, 32))
            If DECO$ = "" Then
               Call MENSERR(24, "Falta Descripción Para el Código Número '" + TRIM$(Str$(COCO%)) + "'.")
               GoTo 10
            End If
         '
         Next U&
      
         CONDI$ = "COD_INTE=" & CI%
         Call ACTUREGISTRO("RELACODPROV", "MARBORRA", CONDI$, 1, REG&) ' HACE UN MARBORRA

         SQLX$ = "SELECT * FROM RELACODPROV "
         SQLX$ = SQLX$ + " WHERE COD_INTE = " & CI%
         Set RS2 = New ADODB.Recordset
         RS2.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
          With RS2
            For I& = 1 To ULTREG&("!PROVIADE") ' RECORRO EL ARCHIVO Y GRABO
              XX$ = REGLEIDO$("!PROVIADE", I&) ' LA INFORMACION DE LA VENTANA
              NPRO% = CVI(Mid$(XX$, 1, 2))
              VIA% = CVI(Mid$(XX$, 3, 2))
              PRECIOX# = CVD(Mid$(XX$, 5, 8))
              MON% = CVI(Mid$(XX$, 13, 2))
              FE% = CVI(Mid$(XX$, 15, 2))
              
              .AddNew   ' YA CONVERTIDO
              
              On Error Resume Next
              !CodiEmpr = CodiEmp%
              On Error GoTo 0
              !NUM_PROV = NPRO%
              !Cod_Inte = CI%
              !VIA = VIA%
              !PRECIO = PRECIOX#
              !MONEDA = MON%
              !FECHA = CVDAT(FE%)
              !MarBorra = 0
              .Update
            Next I&
          End With
         RS2.Close
         Set RS2 = Nothing
         CONDI$ = "MARBORRA = 1"
         Call BORRAREGISTROS("RELACODPROV", CONDI$, REG&) ' BORRO LOS MARCADOS CON 1

         Call TRACENTRAL("PROVIADE", "")
         Call FRESHECHO("PROVIADE")
         Screen.MousePointer = 1
         Call COMUNI("Grabación Finalizada")
End Sub

Private Sub mnuDemora_Click()
      Call TRALOCAL("DEMORAS", "")
      '
10    Screen.MousePointer = 1
      VTB% = VENTABU%("DEMORAS")
      If VTB% < 0 Then Exit Sub
      '
      
      Screen.MousePointer = 11
      For U& = 1 To ULTREG&("!DEMORAS")
         X$ = REGLEIDO$("!DEMORAS", U&)
         COCO% = Asc(Left$(X$, 2))
         If COCO% < 1 Then
            Call MENSERR(24, "Código debe Vía de Entrega Debe ser Numérico.")
            GoTo 10
         End If
         '
      Next U&
      '
      'VALIDO QUE NO SE REPITA EL CODIGO
      For U& = 1 To ULTREG&("!DEMORAS")
         X$ = REGLEIDO$("!DEMORAS", U&)
         COCO% = Asc(Left$(X$, 2))
         For j& = 1 To ULTREG&("!DEMORAS")
            If j& <> U& Then
              X$ = REGLEIDO$("!DEMORAS", j&)
              COCO2% = Asc(Left$(X$, 2))
              If COCO% = COCO2% Then
                Call MENSERR(24, "Código Repetido. (" & COCO2% & ")")
                GoTo 10
              End If
            End If
          Next j&
      Next U&
      
      Call TRACENTRAL("DEMORAS", "")
      Call FRESHECHO("DEMORAS")
      Screen.MousePointer = 1
      Call COMUNI("Grabación Finalizada")
End Sub


Private Sub MNUEDITPRESUP_Click()
    '
10  Call CARGA_PRESUPUESTOS
    Screen.MousePointer = 1
    If RESPUESTA_BUSQUEDA$ = "" Then
      Screen.MousePointer = 1
      Exit Sub
    End If
    '
    NUMERISELEC$ = RESPUESTA_BUSQUEDA$
    'PR-00000063-00
'    SIG_VERSION% = 1 + XVALO(Mid$(NUMERISELEC$, 13, 2)) ' NUEVA VERSION
    NUMERI$ = Left$(NUMERI$, 12) + Right("00" & SIG_VERSION%(NUMERISELEC$), 2) ' CONCATENO AL MISMO PRESUPUESTO
    Text17 = XVALO(Mid$(NUMERISELEC$, 4, 8))
    Text1 = VALOBADA("PRESUVENTA", "NroClie", "NROPRES = " & Text17, "NOMESS")
    Text20 = "" 'para que siempre dispare el evento change
    VERSIOX$ = "00" & SAFETEXT$(SIG_VERSION%(NUMERISELEC$)): VERSIOX$ = Right(VERSIOX$, 2)
    Text20 = VERSIOX$
    Text1.Locked = True
    '
    Call PREPORCOP(NUMERISELEC$)
    Text21.Visible = True

    '
End Sub
Function SIG_VERSION%(PRESUPUESTOX$)
       'PR-00000063-00
       NUMEPRE& = XVALO(Mid$(PRESUPUESTOX$, 4, 8))
       CONDI$ = "NroPres = " & NUMEPRE&
       SIG_VERSION% = 1 + VALOBADA("PRESUVENTA", "MAX(RIGHT(NroPres_Version,2))", CONDI$, "NOMESS")
End Function
Private Sub mnupcopia_Click()

   If XVALO(Text20) > 0 Then Exit Sub ' modo modificacion
   
   'presenta lista de presupuestos
10  Call CARGA_PRESUPUESTOS
    If RESPUESTA_BUSQUEDA$ = "" Then Exit Sub
    NUDOCU$ = RESPUESTA_BUSQUEDA$
   'presento lista de clientes
   
   Call CARGA_INDICE_CLIENTES
   NCLIE% = XVALO(RESPUESTA_BUSQUEDA$)
   If NCLIE% < 1 Then Exit Sub
   '
   Text1 = NCLIE%
   Call PREPORCOP(NUDOCU$)
   '
End Sub
Sub PREPORCOP(NPRESU$)
   '
   NUCLI% = NCLIE% 'CVI(Mid$(XX$, 7, 2))'CLIENTE QUE SE ELIGIO
   Call ABREPRESVEN
   '
   SQLX$ = " SELECT * "
   SQLX$ = SQLX$ + " FROM DETAPRESUPUESTO WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE NroPres_Version = '" & NPRESU$ & "' "
   SQLX$ = SQLX$ + " ORDER BY NuOrdIt "
   '
   Dim PRESUTEMP As ADODB.Recordset
   Set PRESUTEMP = READSET(SQLX$)
    '
   MSF1.Rows = 1
   Screen.MousePointer = 11
   With PRESUTEMP
      Do While Not .EOF
        I& = I& + 1
        CI1% = XVALO(!CodiInt)
        CODE$ = SAFETEXT$(!CodiExt)
        DESCRIX$ = SAFETEXT$(!Descrip)
        PREUNI# = XVALO(!PreUnitItem)
        CANTI = XVALO(!Cantit)
        IMPORTX# = XVALO(!ImToNeItem)
        MSF1.Rows = MSF1.Rows + 1
        MSF1.TextMatrix(I&, 1) = CODE$
        MSF1.TextMatrix(I&, 2) = DESCRIX$
        MSF1.TextMatrix(I&, 3) = TRIM$(FORMATNUM$(CANTI, "F12.1"))
        MSF1.TextMatrix(I&, 5) = TRIM$(FORMATNUM$(PREUNI#, "F12.2"))
        MSF1.TextMatrix(I&, 6) = TRIM$(FORMATNUM$(IMPORTX#, "F12.2"))
      .MoveNext
      Loop
   End With
   PRESUTEMP.Close
   Set PRESUTEMP = Nothing
   '
   Screen.MousePointer = 1
   Call ACTU_TOTAL
   ALPORCOPIA% = 1 'MARCA PARA SABER SI ESTA EN ALTA POR COPIA
End Sub

Private Sub MnuProv_Art_Click()
         '
         Call GENERA_TABLA_RELACION_ART_PROV
5        Call SELECHO("PROVED1")
         '
         NPRO% = XVALO(VALACT1("PROVED1"))
         If NPRO% < 1 Then Exit Sub
         '
         SQLX$ = "SELECT * FROM RELACODPROV  WITH(NOLOCK)"
         SQLX$ = SQLX$ + " WHERE NUM_PROV = " & NPRO% & " AND MARBORRA = 0 "
         Set RS = READSET(SQLX$)
         '
         Call SETULTREG("!ARTPROVD", 0)
         With RS
            On Error Resume Next
            .MoveFirst
            If Err Then
                On Error GoTo 0
                GoTo 10
            End If
            JJ& = 0
             Do While (.EOF = False)
                X$ = REGBLAN$("ARTPROVD")
                Q = !VIA
                Call REPLA(X$, MKI(XVALO(!Cod_Inte)), 1, 2)
                Call REPLA(X$, MKI(XVALO(!VIA)), 3, 2)
                Call REPLA(X$, MKD(XVALO(!PRECIO)), 5, 8)
                Call REPLA(X$, MKI(XVALO(!MONEDA)), 13, 2)
                Call REPLA(X$, MKI(CVINT(!FECHA)), 15, 2)
                JJ& = JJ& + 1
                Call GRAREG("!ARTPROVD", X$, JJ&)
            .MoveNext
            Loop
         End With
         RS.Close
         Set RS = Nothing
         '
10       XYX% = VENTABU%("ARTPROVD")
         If XYX% < 0 Then GoTo 5
         '
         CONDI$ = "NUM_PROV =" & NPRO%
         Call ACTUREGISTRO("RELACODPROV", "MARBORRA", CONDI$, 1, REG&)
    
         
         SQLX$ = "SELECT * FROM RELACODPROV "
         SQLX$ = SQLX$ + " WHERE NUM_PROV = " & NPRO%
         Set RS1 = New ADODB.Recordset
         RS1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
           With RS1
             For I& = 1 To ULTREG&("!ARTPROVD")
               XX$ = REGLEIDO$("!ARTPROVD", I&)
               
               CI% = CVI(Mid$(XX$, 1, 2))
               VIA% = CVI(Mid$(XX$, 3, 2))
               PRECIOX# = CVD(Mid$(XX$, 5, 8))
               MON% = CVI(Mid$(XX$, 13, 2))
               FE% = CVI(Mid$(XX$, 15, 2))
        
               
               CI% = CVI(Mid$(XX$, 1, 2))
               .AddNew   ' YA CONVERTIDO
               On Error Resume Next
               !CodiEmpr = CodiEmp%
               On Error GoTo 0
               !NUM_PROV = NPRO%
               !Cod_Inte = CI%
               !PRECIO = PRECIOX#
               !MONEDA = MON%
               !FECHA = CVDAT(FE%)
               !VIA = VIA%
               !MarBorra = 0
               .Update
             Next I&
           End With
         RS1.Close
         Set RS1 = Nothing
         CONDI$ = CONDI + "AND  MARBORRA = 1  "
         Call BORRAREGISTROS("RELACODPROV", CONDI$, REG&)
         
         Call TRACENTRAL("ARTPROVD", "")
         Call FRESHECHO("ARTPROVD")
         Screen.MousePointer = 1
         Call COMUNI("Grabación Finalizada")

End Sub

Private Sub MnuTabEntr_Click()
      Call TRALOCAL("VIASENTR", "")
      '
10    Screen.MousePointer = 1
      VTB% = VENTABU%("VIASENTR")
      If VTB% < 0 Then Exit Sub
      '
      Screen.MousePointer = 11
      For U& = 1 To ULTREG&("!VIASENTR")
         X$ = REGLEIDO$("!VIASENTR", U&)
         COCO% = Asc(Left$(X$, 1))
         If COCO% < 1 Then
            Call MENSERR(24, "Código debe ser Numérico.")
            GoTo 10
         End If
         '
         DECO$ = TRIM$(Mid$(X$, 2, 32))
         If DECO$ = "" Then
            Call MENSERR(24, "Falta Descripción Para el Código Número '" + TRIM$(Str$(COCO%)) + "'.")
            GoTo 10
         End If
         '
      Next U&
      '
      Call TRACENTRAL("VIASENTR", "")
      Call FRESHECHO("VIASENTR")
       Screen.MousePointer = 1
      Call COMUNI("Grabación Finalizada")
     '
End Sub

Private Sub mnuvisualiza_Click()

10  Call CARGA_PRESUPUESTOS
    If RESPUESTA_BUSQUEDA$ = "" Then Exit Sub
   

    TIDOCUM$ = "Presupuesto"
    NUDOCU$ = RESPUESTA_BUSQUEDA$
    Screen.MousePointer = 11
    '
    For UI& = ULTREG&("PDOCLST") To 1 Step -1
      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(XX$, TIDOCUM$) > 4 Then
        PPXX% = InStr(XX$, "Nro.")
        If PPXX% > 0 Then
          XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
        End If
        POS1% = InStr(XX$, "Presupuesto")
        If POS1% > 0 Then
           NUMECOMPA$ = TRIM$(Mid$(XX$, POS1% + 12))
           If NUMECOMPA$ = TRIM$(NUDOCU$) Then
                DOCUNU& = CVS(Left$(XX$, 4))
                GoTo 5
           End If
        End If
      End If
    Next UI&
    Screen.MousePointer = 1
    MsgBox "DOCMENTO NO ENCONTRADO"
    GoTo 10
   '
5   Call MUESTRADOC(DOCUNU&)
    GoTo 10
End Sub


Private Sub MONEMIS_DblClick()

   Static CAMON%, DMONE$(), TCMONE()
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
           TCMONE(U) = VHX
         Next U&
     End If
   End If
   '
   MONEMIA% = MONEMI%
15 MONEMI% = MONEMI% + 1
   If MONEMI% > CAMON% Then MONEMI% = 1
   If DMONE$(MONEMI%) = "" Then GoTo 15
   '
   MONEMIS = DMONE$(MONEMI%)
   MONEMIS.ForeColor = RGB(0, 0, 255)
   If MONEMI% > 1 Then
      MONEMIS.ForeColor = RGB(255, 0, 0)
   End If
   '
   MONEMIS.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MONEMI%), "F10.4")) + " - Doble-Click para Cambiar Moneda"
   TICAMBIS = FORMATNUM$(TICAMONE(MONEMI%), "F9.4")
   '
   If MONEMI% > 1 Then
     TICAMBIS.Visible = True
     TICAMBIS.Enabled = False
   End If
   '
   Total(0) = FORMATNUM$(XVALO(Total(0)) * TICAMONE(MONEMIA%) / TICAMONE(MONEMI%), "F12.2")
   '
      COEFCON = TICAMONE(MONEMIA%) / TICAMONE(MONEMI%)
      For U& = 1 To MSF1.Rows - 1

        CAN = XVALO(MSF1.TextMatrix(U&, 3))
        'PORDESCUIN = CVS(Mid$(X$, 69, 4))
        COSTOX# = COEFCON * XVALO(MSF1.TextMatrix(U&, 4))
        PRECVENTA# = COEFCON * XVALO(MSF1.TextMatrix(U&, 5))
        IMPORTEX# = COEFCON * XVALO(MSF1.TextMatrix(U&, 6))
        MSF1.TextMatrix(U&, 4) = FORMATNUM$(COSTOX#, "F12.3")
        MSF1.TextMatrix(U&, 5) = FORMATNUM$(PRECVENTA#, "F12.3")
        MSF1.TextMatrix(U&, 6) = FORMATNUM$(IMPORTEX#, "F12.2")
        
      Next U&
      '
      Call ACTU_TOTAL
'      Call CARLISPRES
'      Call CIERRARCH("!CARDETPR")
'      Call CALTOPRES
'   End If
   '
End Sub




Private Sub MSF1_DblClick()
   If MSF1.Row >= 1 Then
     FRMEDIT_PRESU.Show 1
   End If

End Sub

Private Sub MSFlexGrid2_Click()

End Sub

Private Sub PORDESCU_Change()
   Call ACTU_TOTAL
End Sub

Private Sub Text1_Change()
   
   Call LIMPIAR_TEXT_DETALLE
   Call LIMPIARPRESU
   MSF1.Rows = 1
   
   If TRIM$(Text1.TEXT) <> "" Then
      Call CARGADATOS_CLIENTE
   End If
   NC% = XVALO(Text1)
   'VALIDO SI LA LISTA ACTIVA ES VALIDA
   If XVALO(Text1) > 0 Then
    NLIST% = XVALO(Text23)
'    If NLIST% < 1 Or TRIM$(Text22) = "" Then
'      TX$ = "El CLIENTE ACTIVO NO POSEE LISTA DE PRECIOS ASIGNADA" + vbCrLf
'      TX$ = TX$ + "POR DEFECTO SE ASIGNA LISTA 1"
'      MsgBox TX$
'      NLIST% = 1
'      Text23 = NLIST%
'    End If
   End If
   Label26 = CODICLI$(NC%)
   'NUMERO DE PRESUPUESTO
'   If Text21.Visible = False Then
'    Text17 = 1 + MAXNUMCOD&("PRESUVENTA", "NroPres")
'    If XVALO(Text20) < 1 Then Text20 = "-00"
'    Text21.Visible = False
'
'   End If
   '
   'FECHA
   DTPicker1.Value = Date
   
'   If Command1.Enabled = True Then
'       Call FRMPRESUAHP.Command1_Click
'   End If
End Sub

Sub LIMPIARPRESU() 'SE LE PUEDE PASAR UN TEXTBOX QUE ESTE NO SE LIMPIE
    
    Text2 = ""
    Text3 = ""
    Text4 = ""
    Text5 = ""
    Text8 = ""
    Text18 = ""
    TEXT13 = ""
    Text14 = ""
    Text19 = ""
    Text23 = ""
    Text20 = ""
    '
End Sub
Sub CARGADATOS_CLIENTE()
   CUENTA& = XVALO(Text1)
   If CUENTA& < 1 Then: Call LIMPIARPRESU: Exit Sub
   '
   CONDI$ = "Nume_Cli = " & CUENTA&
   '
   'VALIDO SI EXISTE
   CNRG& = CantRegistros&("DEUDOR12", CONDI$)
   If CNRG& < 1 Then Call LIMPIARPRESU: Exit Sub
   '
   'CARGO DATOS DE CLIENTE
   Text2 = VALOBADA("DEUDOR12", "Raso_Cli", CONDI$, "NOMESS")
   Text3 = VALOBADA("DEUDOR12", "Domi_Cli", CONDI$, "NOMESS")
   Text4 = VALOBADA("DEUDOR12", "LOCA_CLI", CONDI$, "NOMESS")
   Text8 = VALOBADA("DEUDOR12", "Tele_Cli", CONDI$, "NOMESS")
   Text5 = VALOBADA("DEUDOR12", "CELU_CLI", CONDI$, "NOMESS")
   Text18 = VALOBADA("DEUDOR12", "Mail_Cli", CONDI$, "NOMESS")
   Text19 = VALOBADA("DEUDOR12", "Nota_Cli", CONDI$, "NOMESS")
   Text23 = CEROIGUALVACIO(XVALO(VALOBADA("DEUDOR12", "Lpre_Cli", CONDI$, "NOMESS")))
   '

End Sub

Private Sub Text20_Change()
  If XVALO(Text20) > 0 Then
     Text21.Visible = True
     'Call APLICAR_SKIN(Me, App.Path & "\SKINS\GOLDEN.skn")
  Else
     'Call APLICAR_SKIN(Me, App.Path & "\SKINS\dogmax.skn")
     Text21.Visible = False
  End If
End Sub


Private Sub Text23_Change()
'   CODI% = XVALO(Text23)
'   Text22 = ECOLPRECIOS$(CODI%)

End Sub

Private Sub TXTCANTIDAD_Change()

   Resultado# = XVALO(TXTCANTIDAD) * XVALO(TXTPREUNIT)
   TXTIMPORTE = FORMATNUM$(Resultado#, "F14.2")

End Sub

Private Sub TXTCANTIDAD_KeyPress(KeyAscii As Integer)
   
    If KeyAscii = 27 Then
        KeyAscii = 0
        Call LIMPIAR_TEXT_DETALLE
    End If
   
'    If KeyAscii = 13 Then TXTPREUNIT.SetFocus
   
' *********** CODIGO ANTERIOR ************

    If KeyAscii = 27 Then
        KeyAscii = 0
        Call LIMPIAR_TEXT_DETALLE
    End If


    If KeyAscii = 13 Then
        KeyAscii = 0
        '
        'VALIDO EL CODIGO
        CI1% = CODINT%(txtcodigo)
        If CI1% < 1 Then
           MsgBox "CÓDIGO INEXISTENTE O NO VÁLIDO"
           Exit Sub
        End If

        'VALIDO QUE TENGA INGRESADA CANTIDAD
        If XVALO(TXTCANTIDAD.TEXT) < 0.05 Then
           MsgBox "FALTA CANTIDAD"
           TXTCANTIDAD.SetFocus
           Exit Sub
        End If
        'VALIDO QUE TENGA PRECIO
        If XVALO(TXTPREUNIT.TEXT) < 0.05 Then
           KeyAscii = 0
           MsgBox "FALTA PRECIO UNITARIO"
           TXTPREUNIT.SetFocus
           Exit Sub
        End If
        '
'        If VALIDA_COD_GRID(txtcodigo, FRMPRESU.MSF1, 1) > 0 Then
'         MsgBox "MODIFIQUE DATOS SOBRE EL MISMO"
'         GoTo 50
'        End If

        TXTPREUNIT.BackColor = &HE0E0E0 'CAMBIA A GRIS
        TXTPREUNIT.Locked = True 'DESHABILITO
        cargar_datos
        txtcodigo = ""
        txtcodigo.SetFocus
50      Call LIMPIAR_TEXT_DETALLE
          
        If Command1.Enabled = True Then
          
          Call FRMPRESUAHP.Command1_Click
          
        End If
    End If


    

End Sub
Private Sub cargar_datos()
    '
    MSF1.Cols = 7
    MSF1.Rows = MSF1.Rows + 1
    MSF1.Row = MSF1.Rows - 1
    CI1% = CODINT%(txtcodigo.TEXT)
    MSF1.TextMatrix(MSF1.Rows - 1, 1) = txtcodigo.TEXT
    MSF1.TextMatrix(MSF1.Rows - 1, 2) = DESCRIT0$(CI1%)
    MSF1.TextMatrix(MSF1.Rows - 1, 3) = FORMATNUM$(TXTCANTIDAD.TEXT, "F10.1")
    
    MSF1.TextMatrix(MSF1.Rows - 1, 5) = FORMATNUM$(TXTPREUNIT.TEXT, "f12.3")
    MSF1.TextMatrix(MSF1.Rows - 1, 6) = FORMATNUM$(TXTIMPORTE.TEXT, "f14.2")
    
'    MSF1.TextMatrix(MSF1.Rows - 1, 6) = "0" ' INDICADOR DE CODIGO OPCIONAL
'    If Check1.Value = 1 Then
'        MSF1.TextMatrix(MSF1.Rows - 1, 6) = "1"
'        MSF1.TextMatrix(MSF1.Rows - 1, 5) = ""
'    End If
'    MSF1.ColWidth(6) = 0 ' OCULTA LA COLUMNA
'    Check1.Value = 0 ' SETEA NO OPCIONAL POR DEFAULT
    '
    ' ORDENA GRILLA POR COLUMNA 6 Y EN FORMA ASCENDENTE (VALOR 7 DE SORT DATOS DE TIPO STRING)
    MSF1.Col = 6
    MSF1.Sort = 7
    '
    'SUMO EL IMPORTE
    Call ACTU_TOTAL
    '
    On Error Resume Next
    VISIROWS = (MSF1.Height / MSF1.RowHeight(j&)) - 4
    If Err > 0 Then
      VISIROWS = 1
    End If
    On Error GoTo 0
    If VISIROWS > 0 Then
       If MSF1.TopRow < MSF1.Row - VISIROWS Then
          MSF1.TopRow = MSF1.Row - VISIROWS
       End If
    End If
    
End Sub
Sub ACTU_TOTAL()


   '
   Static CLIA$, RNI%, APLIVA%
   If Text1.TEXT <> CLIA$ Then
     CLIA$ = Text1.TEXT: NCLIE% = XVALO(CLIA$)
     RNI% = 0: APLIVA% = 1
     If NCLIE% > 0 Then
       PVCLI% = PIVACLI%(NCLIE%)
       If PVCLI% = 2 Then RNI% = 1
       If PVCLI% = 5 Or PVCLI% = 7 Then APLIVA% = 0
     End If
   End If
   '
   SUTOTPED# = 0: SUIVA# = 0: SUIVARNI# = 0: TOIVA# = 0: TOTOTA# = 0
   For U& = 1 To MSF1.Rows - 1
     CIXXI% = CODINT%(MSF1.TextMatrix(U&, 1))
     CAN = XVALO(MSF1.TextMatrix(U&, 3))
'     PORDESCUIN = CVS(Mid$(X$, 69, 4))
'     If PORDESCUEIN < 0 Or PORDESCUEIN > 99 Then PORDESCUEIN = 0
     If CAN > 0.05 Then
       PRELI = XVALO(MSF1.TextMatrix(U&, 5))
       'If PRELI < 0.01 Then
       '  PRELI = PRELISTA(ARCHIPRE$, CIXXI%)
       'End If
       PREUNI = PRELI * (100 - PORDESCUIN) / 100
       TOT# = CDbl(CAN * PREUNI)
       SUTOTPED# = SUTOTPED# + TOT#
       SUIVA# = SUIVA# + (Int(TOT# * PORCEIVA(CIXXI%) + 0.5)) / 100
'       Call REPLA(X$, MKS$(PREUNI) + MKD$(TOT#), 51, 12)
'       Call GRAREG("!CARDETPR", X$, U&)
     End If
   Next U&
   '
   Total(0).TEXT = FORMATNUM$(SUTOTPED#, "F12.2")
   Total(0).Refresh
   '
   DESC$ = TRIM$(PORDESCU.TEXT)
   PODES = COEFDESCU(DESC$)
   TTTX# = SUNEFA2# * (100 - PODES) / 100
3   TTT1 = XVALO(Total(0).TEXT) * PODES / 100
   Total(1).TEXT = CSTRING$(MKS$(TTT1), 3, 11, 2, 2)
   '
   TTT3 = XVALO(Total(0).TEXT) - XVALO(Total(1).TEXT)
   '
    ACUMU_IMPORTE# = 0
'    For I& = 1 To MSF1.Rows - 1
'        IMPORTE# = XVALO(MSF1.TextMatrix(I&, 5))
'        ACUMU_IMPORTE# = ACUMU_IMPORTE# + IMPORTE#
'    Next I&
    TTTX# = SUIVA# * (100 - PODES) / 100
    Text33 = FORMATNUM$(XVALO(Total(0)) - XVALO(Total(1)), "F12.2")
    Text31 = FORMATNUM$(TTTX#, "f12.2")
    Text14 = FORMATNUM$(XVALO(Text33) + XVALO(Text31), "F12.2")

End Sub
Private Sub txtcodigo_Change()


    TXTDESCRIPCION = ""
    TXTCANTIDAD = ""
    TXTPREUNIT = ""
    TXTIMPORTE = ""
    'DESHABILITO EL TEXT DE PRECIO UNITARIO
    TXTPREUNIT.BackColor = &HE0E0E0 'CAMBIA A GRIS
    TXTPREUNIT.Locked = True 'DESHABILITO
    MSF3.Rows = 1
    TXTPREUNIT.ToolTipText = ""

    If TRIM$(txtcodigo) = "" Then Exit Sub
    'VALIDO SI LA LISTA ACTIVA ES VALIDA
    NLIST% = XVALO(Text23)
    If NLIST% < 1 Or TRIM$(Text22) = "" Then
      NLIST% = 1
    End If
    '
    'BUSCO LOS DATOS
    CI1% = CODINT%(txtcodigo)
    If CI1% < 1 Then
       Exit Sub
    End If
    
'    CONDI$ = " CODINT = " & CI1%
    TXTDESCRIPCION = DESCRIT0(CI1%)
    TXTCANTIDAD = ""
    TXTPREUNIT = FORMATNUM$(PRELISTA(TRIM$(Str$(NLIST%)), CI1%), "F12.2")
    TXTIMPORTE = ""
    ' SI SE SELECCIONO DE LA LISTA PASA EL FOCO A LA CANTIDAD
    If Command1.Enabled = False Then
       TXTCANTIDAD.SetFocus
    End If
  
    
    
End Sub



Private Sub txtcodigo_GotFocus()
'   If TRIM$(txtcodigo.TEXT) <> "" Then
'       Command1_Click
'   End If
End Sub


Private Sub txtcodigo_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then TXTCANTIDAD.SetFocus
End Sub

Private Sub TXTDESCRIPCION_Change()
    MSF3.Rows = 1
    If TRIM$(TXTDESCRIPCION.TEXT) <> "" Then
       CI1& = CODINT(txtcodigo.TEXT)
       If CI1& > 0 Then
          Call ULTIMOS3PRESUPUESTOS(FRMPRESUAHP.MSF3, CI1&)
       End If
    End If
End Sub


Private Sub TXTPREUNIT_Change()

   Resultado# = XVALO(TXTCANTIDAD) * XVALO(TXTPREUNIT)
   TXTIMPORTE = FORMATNUM$(Resultado#, "F14.2")

End Sub
Sub LIMPIAR_TEXT_DETALLE()
    TXTDESCRIPCION = ""
    TXTCANTIDAD = ""
    TXTPREUNIT = ""
    TXTIMPORTE = ""
    txtcodigo = ""
    'DESHABILITO EL TEXT DE PRECIO UNITARIO
    TXTPREUNIT.Locked = False
    TXTPREUNIT.ToolTipText = ""
    Text35 = ""
End Sub

Sub TXTPREUNIT_DblClick()

    If TXTPREUNIT.Locked = False Then
       TXTPREUNIT.BackColor = &HE0E0E0
       TXTPREUNIT.Locked = True
    Else
       TXTPREUNIT.Locked = False
       TXTPREUNIT.BackColor = &H80000005
       Call PINTATEXT(TXTPREUNIT)
    End If
    
End Sub

Private Sub TXTPREUNIT_KeyPress(KeyAscii As Integer)


    If TXTPREUNIT.Locked = True Then Exit Sub
    '
    
    If KeyAscii = 27 Then
        KeyAscii = 0
        Call LIMPIAR_TEXT_DETALLE
    End If
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If XVALO(TXTCANTIDAD.TEXT) < 0.05 Then
           KeyAscii = 0
           MsgBox "FALTA CANTIDAD"
           TXTCANTIDAD.SetFocus
           Exit Sub
        End If
        
        If XVALO(TXTPREUNIT.TEXT) < 0.05 Then
           KeyAscii = 0
           MsgBox "FALTA PRECIO UNITARIO"
           TXTPREUNIT.SetFocus
           Exit Sub
        End If

'        If VALIDA_COD_GRID(txtcodigo, FRMPRESU.MSF1, 1) > 0 Then
'         MsgBox "MODIFIQUE DATOS SOBRE EL MISMO"
'         GoTo 50
'        End If
        TXTPREUNIT.BackColor = &HE0E0E0 'CAMBIA A GRIS
        TXTPREUNIT.Locked = True 'DESHABILITO

        cargar_datos
        txtcodigo = ""
        txtcodigo.SetFocus
50      Call LIMPIAR_TEXT_DETALLE
        If Command1.Enabled = True Then
           Call FRMPRESUAHP.Command1_Click
        End If

     End If
    
End Sub
   
Private Sub TXTPREUNIT_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode 'ESTO ES PARA QUE DESP
        Case 115 'pulso  f4
        Call TXTPREUNIT_DblClick

    End Select

End Sub
