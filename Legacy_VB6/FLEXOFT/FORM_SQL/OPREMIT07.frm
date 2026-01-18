VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form OPREMIT07 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración y Formulario de Remito de Expedición"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11655
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
   ScaleHeight     =   6045
   ScaleWidth      =   11655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture999 
      Height          =   330
      Index           =   0
      Left            =   0
      ScaleHeight     =   270
      ScaleWidth      =   1110
      TabIndex        =   50
      Top             =   240
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   6045
      Left            =   10800
      ScaleHeight     =   6015
      ScaleWidth      =   1350
      TabIndex        =   34
      Top             =   0
      Width           =   1380
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   105
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   72
         Top             =   5520
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   73
            Top             =   -120
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Direct"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   105
         Picture         =   "OPREMIT07.frx":0000
         TabIndex        =   52
         ToolTipText     =   "Accesos Directos"
         Top             =   3060
         Width           =   650
      End
      Begin VB.CommandButton Command19 
         Caption         =   "Equiv"
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
         Picture         =   "OPREMIT07.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Códigos Equivalentes de Clientes"
         Top             =   2370
         Width           =   650
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
         Height          =   690
         Left            =   105
         Picture         =   "OPREMIT07.frx":0294
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Emite Reporte Solicitado"
         Top             =   1680
         Width           =   650
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Help"
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
         Picture         =   "OPREMIT07.frx":059E
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   890
         Width           =   650
      End
      Begin VB.CommandButton Command1 
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
         Height          =   690
         Left            =   105
         Picture         =   "OPREMIT07.frx":08A8
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Abandona la Aplicación"
         Top             =   200
         Width           =   650
      End
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00CDDADA&
      Caption         =   "Opciones Avanzadas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5925
      Left            =   5400
      TabIndex        =   20
      Top             =   80
      Width           =   5265
      Begin VB.Frame Frame12 
         BackColor       =   &H00CDDADA&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   470
         Index           =   4
         Left            =   240
         TabIndex        =   84
         Top             =   4920
         Width           =   4905
         Begin VB.CheckBox Check7 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Lote Imprime Solo en Parametro NRO-LOTE"
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
            Height          =   300
            Left            =   120
            TabIndex        =   85
            ToolTipText     =   "Imprime Observaciones x Items de Pedido Cargadas al realizar el mismo."
            Top             =   130
            Width           =   4005
         End
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00CDDADA&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   470
         Index           =   3
         Left            =   240
         TabIndex        =   74
         Top             =   5340
         Width           =   4905
         Begin VB.CheckBox Check6 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Muestra Datos de Entrega del Pedido (Solo si hay 1 pedido activo)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   120
            TabIndex        =   75
            ToolTipText     =   "Imprime Observaciones x Items de Pedido Cargadas al realizar el mismo."
            Top             =   120
            Width           =   4725
         End
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00CDDADA&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   470
         Index           =   2
         Left            =   200
         TabIndex        =   70
         Top             =   1780
         Width           =   4905
         Begin VB.CheckBox Check5 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Imprime Observaciones por Items de Pedido"
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
            Height          =   300
            Left            =   120
            TabIndex        =   71
            ToolTipText     =   "Imprime Observaciones x Items de Pedido Cargadas al realizar el mismo."
            Top             =   130
            Width           =   4605
         End
      End
      Begin VB.Frame Frame16 
         BackColor       =   &H00CDDADA&
         Caption         =   "PUERTO DE IMPRESION DE REMITOS C/ENTREGA EDITADA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   200
         TabIndex        =   67
         Top             =   3650
         Width           =   4905
         Begin VB.ComboBox Combo2 
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
            Height          =   345
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   68
            Top             =   280
            Width           =   3555
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Impresora"
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
            Index           =   3
            Left            =   120
            TabIndex        =   69
            Top             =   420
            Width           =   930
         End
      End
      Begin VB.Frame Frame15 
         BackColor       =   &H00CDDADA&
         Caption         =   "PUERTO DE IMPRESION ETIQUETAS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   200
         TabIndex        =   57
         Top             =   2900
         Width           =   4905
         Begin VB.ComboBox Combo1 
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
            Height          =   345
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   59
            Top             =   280
            Width           =   3555
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Impresora"
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
            Index           =   2
            Left            =   120
            TabIndex        =   58
            Top             =   420
            Width           =   930
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00CDDADA&
         Caption         =   "IDENTIFICACION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   3650
         TabIndex        =   43
         Top             =   180
         Width           =   1450
         Begin VB.CheckBox Check3 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Número    de Serie"
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
            Height          =   400
            Left            =   120
            TabIndex        =   44
            Top             =   190
            Width           =   1125
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00CDDADA&
         Caption         =   "PUERTO DE IMPRESION - CANTIDAD DE COPIAS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   680
         Left            =   200
         TabIndex        =   38
         Top             =   2230
         Width           =   4905
         Begin VB.TextBox Text10 
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
            Height          =   330
            Left            =   4305
            TabIndex        =   40
            Top             =   250
            Width           =   330
         End
         Begin VB.TextBox Text12 
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
            Height          =   330
            Left            =   1155
            TabIndex        =   39
            Top             =   250
            Width           =   2955
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Remitos"
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
            Index           =   5
            Left            =   -105
            TabIndex        =   41
            Top             =   390
            Width           =   1170
         End
      End
      Begin VB.Frame Frame14 
         BackColor       =   &H00CDDADA&
         Caption         =   "CODIGO-CLIENTE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   1620
         TabIndex        =   33
         Top             =   180
         Width           =   2000
         Begin VB.CheckBox Check2 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Reemplaza por Código Clente"
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
            Height          =   400
            Left            =   120
            TabIndex        =   61
            Top             =   190
            Width           =   1485
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00CDDADA&
         Caption         =   "STOCKS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   200
         TabIndex        =   26
         Top             =   180
         Width           =   1395
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Verificar Existencia "
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
            Height          =   400
            Left            =   120
            TabIndex        =   27
            Top             =   190
            Width           =   1155
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00CDDADA&
         Caption         =   "EXPEDICION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   940
         Left            =   200
         TabIndex        =   21
         Top             =   840
         Width           =   4905
         Begin VB.TextBox Text9 
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
            Left            =   2070
            TabIndex        =   30
            TabStop         =   0   'False
            Text            =   " "
            Top             =   550
            Width           =   2730
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
            Left            =   1350
            TabIndex        =   29
            Text            =   " "
            Top             =   550
            Width           =   540
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
            Left            =   1875
            TabIndex        =   28
            Top             =   550
            Width           =   175
         End
         Begin VB.TextBox Text8 
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
            Height          =   315
            Left            =   2070
            TabIndex        =   24
            TabStop         =   0   'False
            Text            =   " "
            Top             =   185
            Width           =   2730
         End
         Begin VB.TextBox Text7 
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
            Height          =   315
            Left            =   1350
            TabIndex        =   23
            Text            =   " "
            Top             =   185
            Width           =   540
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
            Height          =   300
            Left            =   1875
            TabIndex        =   22
            Top             =   185
            Width           =   175
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "L. de Precios"
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
            Index           =   0
            Left            =   280
            TabIndex        =   31
            Top             =   640
            Width           =   1230
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Dep. Salida"
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
            Left            =   440
            TabIndex        =   25
            Top             =   280
            Width           =   1110
         End
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00CDDADA&
         Caption         =   "M3 POR PALLET (LISTA DE EMPAQUE)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   200
         TabIndex        =   47
         Top             =   4380
         Width           =   4905
         Begin VB.TextBox Text11 
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
            Left            =   3600
            TabIndex        =   48
            Top             =   140
            Width           =   1155
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "LIMITE M3"
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
            Index           =   1
            Left            =   2050
            TabIndex        =   49
            Top             =   250
            Width           =   1410
         End
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00CDDADA&
      Caption         =   "REMITO DE EXPEDICION"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5910
      Left            =   120
      TabIndex        =   0
      Top             =   80
      Width           =   5160
      Begin VB.Frame Frame17 
         BackColor       =   &H00CDDADA&
         Caption         =   "Pref-Remito S/Cargo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Left            =   120
         TabIndex        =   76
         Top             =   4080
         Width           =   4875
         Begin VB.CommandButton Command13 
            Caption         =   "Configuración de Pre-Fijo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   120
            TabIndex        =   78
            Top             =   220
            Width           =   4650
         End
         Begin VB.CommandButton Command15 
            Caption         =   "Abrir"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   650
            Left            =   3960
            Picture         =   "OPREMIT07.frx":09F2
            Style           =   1  'Graphical
            TabIndex        =   82
            Top             =   240
            Width           =   765
         End
         Begin VB.TextBox Text18 
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
            TabIndex        =   81
            TabStop         =   0   'False
            Text            =   " "
            Top             =   600
            Width           =   3735
         End
         Begin VB.CommandButton Command14 
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
            Left            =   3670
            TabIndex        =   80
            Top             =   240
            Width           =   175
         End
         Begin VB.TextBox Text17 
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
            Height          =   330
            Left            =   2520
            TabIndex        =   79
            Text            =   " "
            Top             =   240
            Width           =   1140
         End
         Begin VB.TextBox Text16 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   720
            TabIndex        =   77
            TabStop         =   0   'False
            Text            =   "0000"
            Top             =   230
            Width           =   615
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Pre-Fijo:"
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
            Index           =   4
            Left            =   120
            TabIndex        =   83
            Top             =   360
            Width           =   570
         End
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00CDDADA&
         Caption         =   "UTILIZA TABLA DE SALDOS DE LOTE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   120
         TabIndex        =   64
         Top             =   5160
         Width           =   4905
         Begin VB.CommandButton Command12 
            Caption         =   "Configuración Solo Supervisada"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   300
            TabIndex        =   65
            Top             =   220
            Width           =   4245
         End
         Begin VB.CheckBox Check4 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Utiliza Consumo x lotes Tabla salotedep"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   840
            TabIndex        =   66
            Top             =   240
            Width           =   3645
         End
      End
      Begin VB.TextBox Text15 
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
         Height          =   330
         Left            =   2325
         TabIndex        =   62
         Top             =   1680
         Width           =   1550
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Abrir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   4245
         Picture         =   "OPREMIT07.frx":0CFC
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   950
         Width           =   765
      End
      Begin VB.TextBox Text14 
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
         TabIndex        =   55
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1320
         Width           =   3750
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
         Height          =   360
         Left            =   3720
         TabIndex        =   54
         Top             =   950
         Width           =   175
      End
      Begin VB.TextBox Text13 
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
         Height          =   330
         Left            =   2325
         TabIndex        =   53
         Text            =   " "
         Top             =   950
         Width           =   1380
      End
      Begin VB.Frame Frame18 
         BackColor       =   &H00CDDADA&
         Caption         =   "LIMITE CARGA (U$)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   2955
         TabIndex        =   45
         Top             =   3360
         Width           =   2055
         Begin VB.TextBox Text23 
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
            Height          =   330
            Left            =   120
            TabIndex        =   46
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00CDDADA&
         Caption         =   "PREFIJO"
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
         Left            =   3720
         TabIndex        =   18
         Top             =   2640
         Width           =   1275
         Begin VB.TextBox Text4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   210
            TabIndex        =   19
            TabStop         =   0   'False
            Text            =   "0001"
            Top             =   230
            Width           =   870
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00CDDADA&
         Caption         =   "PAPEL"
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
         Left            =   120
         TabIndex        =   15
         Top             =   2640
         Width           =   3450
         Begin VB.OptionButton Option4 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Hoja Suelta"
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
            Height          =   480
            Left            =   1785
            TabIndex        =   17
            Top             =   180
            Width           =   930
         End
         Begin VB.OptionButton Option3 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Formulario Continuo"
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
            Height          =   360
            Left            =   315
            TabIndex        =   16
            Top             =   240
            Value           =   -1  'True
            Width           =   1635
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00CDDADA&
         Caption         =   "TALONARIO EN USO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   120
         TabIndex        =   11
         Top             =   3360
         Width           =   2685
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
            Height          =   330
            Left            =   2445
            TabIndex        =   13
            Top             =   240
            Width           =   175
         End
         Begin VB.TextBox Text3 
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
            Height          =   330
            Left            =   1320
            TabIndex        =   12
            Text            =   " "
            Top             =   240
            Width           =   1185
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Nueva Numeración"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   0
            TabIndex        =   14
            Top             =   240
            Width           =   1245
         End
      End
      Begin VB.TextBox Text1 
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
         Height          =   330
         Left            =   2325
         TabIndex        =   10
         Text            =   " "
         Top             =   200
         Width           =   1380
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
         Height          =   360
         Left            =   3720
         TabIndex        =   9
         Top             =   200
         Width           =   175
      End
      Begin VB.TextBox Text2 
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
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   " "
         Top             =   550
         Width           =   3750
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Abrir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   4245
         Picture         =   "OPREMIT07.frx":1006
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   200
         Width           =   765
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00CDDADA&
         Caption         =   "NUMERACION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   120
         TabIndex        =   3
         Top             =   1980
         Width           =   3750
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Pre-Impresa"
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
            Height          =   240
            Left            =   315
            TabIndex        =   5
            Top             =   300
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.OptionButton Option2 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Auto-Numeración"
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
            Height          =   240
            Left            =   1785
            TabIndex        =   4
            Top             =   300
            Width           =   1800
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00CDDADA&
         Caption         =   "C.MOV"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   3960
         TabIndex        =   1
         Top             =   1980
         Width           =   855
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
            Height          =   345
            Left            =   130
            TabIndex        =   2
            TabStop         =   0   'False
            Text            =   "RT"
            Top             =   240
            Width           =   450
         End
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Ventana de Detalle"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   840
         TabIndex        =   63
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label5 
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
         Left            =   480
         TabIndex        =   60
         Top             =   1070
         Width           =   1725
      End
      Begin VB.Label Label2 
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
         Left            =   480
         TabIndex        =   6
         Top             =   330
         Width           =   1725
      End
   End
   Begin VB.Frame Frame13 
      BackColor       =   &H00CDDADA&
      BorderStyle     =   0  'None
      Caption         =   "Frame13"
      Height          =   1035
      Left            =   4410
      TabIndex        =   32
      Top             =   105
      Width           =   1140
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "OPREMIT07.frx":1310
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "OPREMIT07.frx":1544
      TabIndex        =   51
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "OPREMIT07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call COMUNI("Re-Configuración Completada.\  \Re-Ingrese Ahora a la Aplicación\de Despachos por Remito.")
    Call FINAL("")
End Sub

Private Sub Command10_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text13.TEXT = VDEVU$
        Text14.TEXT = VALACT2$("INDIFRM")
    End If

End Sub

Private Sub Command11_Click()

    Call CONECRUN("FLFORMS/" + TRIM$(Text13.TEXT), "Formularios de Impresion")
    If EXISTE%("FLFORMS.EXE") < 1 Then
        Call FINAL("")
    End If

End Sub



Private Sub Command12_Click()
  If CLACONTRA% < 1 Then
     Exit Sub
   End If
   '
   Command12.Visible = False
End Sub

Private Sub Command13_Click()
   Command13.Enabled = False
   If Command13.Caption = "Grabar" Then
      PREFIREM$ = TRIM$(Control$("", "PREFIREM")): If PREFIREM$ = "" Then PREFIREM$ = "0001"
      If FORMATOCON0$(XVALO(PREFIREM$), 4) = FORMATOCON0$(XVALO(Text16.TEXT), 4) Then
         Call COMUNI("Imposible Grabar Mismo Pre-Fijo Para Remito Sin Cargo y Remito Normal")
         GoTo 99
      End If
      Call GRACONTROL("", "PREFSCAR", FORMATOCON0$(XVALO(Text16.TEXT), 4))
      Call GRACONTROL$("", "FORREMSC", Text17.TEXT)
      Command13.Caption = "Configurar Pre-Fijo " & XVALO(Text16.TEXT)
      If TRIM$(Text17) <> "" Then
         Command13.Caption = Command13.Caption & " - FRM:" & Text17.TEXT
      End If
      
      Command13.Height = 735
      Command13.Left = 120
      Command13.Top = 220
      Command13.Width = 4650
      
      GoTo 99
   End If
   
   If CLACONTRA% > 0 Then
      Command13.Height = 375
      Command13.Left = 1400
      Command13.Top = 240
      Command13.Width = 1095
      Command13.Caption = "Grabar"
   End If
   
99 Command13.Enabled = True
End Sub

Private Sub Command13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If Button = 2 Then
       TX1$ = "Si se Configura un Prefijo de Remito para que se Marque el  Mismo sin Cargo" & vbCrLf
       TX1$ = TX1$ + "Cuando se Grabe  Presentará un Mensaje Dando  Opción" & vbCrLf
       TX1$ = TX1$ + "de Marcarlo Como sin Cargo" & vbCrLf
       TX1$ = TX1$ + "En Caso Afirmativo se Guardará con el Prefijo Configurado" & vbCrLf
       TX1$ = TX1$ + "Pero con La Marca S/Cargo" & vbCrLf
       MsgBox TX1$
     End If

End Sub


Private Sub Command14_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text17.TEXT = VDEVU$
        Text18.TEXT = VALACT2$("INDIFRM")
    End If

End Sub

Private Sub Command15_Click()
    Call CONECRUN("FLFORMS/" + TRIM$(Text17.TEXT), "Formularios de Impresion")
    If EXISTE%("FLFORMS.EXE") < 1 Then
        Call FINAL("")
    End If

End Sub

'\\\OT-05-0241-WAR-02/05/05///
Private Sub Command19_Click()
   Command19.Enabled = False
   OPX% = COMALTER("Opciones de Trabajo\\Modificar Tabla\Listar Codigos Cliente")
   If OPX% = 1 Then
      Call FINAL("*CODORCL")
   ElseIf OPX% = 2 Then
      Call FILESORT("REPLACOC", "", 1, 1, "ASC")
      Call FINAL("*REPLACOC")
   End If
   Command19.Enabled = True
End Sub
'\\\OT-05-0241-WAR-FIN///

Private Sub Command2_Click()
    '
    FENUE$ = TRIM$(Text3.TEXT)
    If FENUE$ = "" Then
       FENUE$ = "01/01/80"
    End If
    If FECHANUM(FENUE$) < 1 Then
       Call MENSERR(24, "Fecha no Válida")
       Text3.SetFocus
       Exit Sub
    End If
    '
    CMO$ = UCase$(TRIM$(Text5.TEXT))
    If Len(CMO$) <> 2 Or CMO$ = "RF" Then
       Call MENSERR(24, "Código de Movimiento no Válido")
       Text5.TEXT = "RT"
       Text5.SetFocus
       Exit Sub
    End If
    '
    Call GRACONTROL("", "FOREMIFA", Text1.TEXT)
    Call GRACONTROL("", "FORETIRE", Text13.TEXT)
    Call GRACONTROL("", "DEPOEXPE", Text7.TEXT)
    Call GRACONTROL$("", "FORREMSC", Text17.TEXT)
    Call GRACONTROL$("", "LOTEPARA", Check7.Value)
    
    MONUME$ = "NO"
    If Check6.Value = 1 Then MONUME$ = "SI"
    Call GRACONTROL$("REMSTAN", "DATENTPE", MONUME$) 'MOSTRAR DATOS DE ENTREGA DEL PEDIDO SOLO SI HAY UN PEDIDO
    
    LIPREXPE$ = ""
    If Val(Text6.TEXT) > 0 Then
      If Val(Text6.TEXT) < 256 Then
        If ECOARCH$("LISTAS", Chr$(Val(Text6.TEXT))) <> "" Then
          LIPREXPE$ = TRIM$(Str$(Val(Text6.TEXT)))
        End If
      End If
    End If
    Call GRACONTROL("", "LIPREXPE", LIPREXPE$)
    '
    MONUME$ = "AUTONUME": If Option1.Value = True Then MONUME$ = "PREIMPRE"
    Call GRACONTROL("", "NUREMIFA", MONUME$)
    '
    TIPAPEL$ = "FORMUCON": If Option4.Value = True Then TIPAPEL$ = "HOJASUEL"
    Call GRACONTROL("", "TIPAREMI", TIPAPEL$)
    '
    MODOEXPE$ = "REMIFAC": TEXPE$ = "Remito de Ventas"
    FOREMIT07.Caption = "Despacho de Materiales - " + TEXPE$
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        FOREMIT07.Caption = FOREMIT07.Caption + "  -  " + EPAC$
    End If
    Call GRACONTROL("", "MODOEXPE", MODOEXPE$)
    '
    'CONTROL DE CAMBIO DE FECHA DE NUEVA NUMERACION PARA TRIGGER DIS_MAXDOPRINUM_insertar
    FECHAGUARDADA% = FECHANUM(Control$("", "FENUREFA")) 'LEO ANTES DE GUARDAR
    FECHAEDITADA% = FECHANUM(FENUE$) ' FENUE$ LO LEE ARRIBA
    If FECHAGUARDADA% <> FECHAEDITADA% Then
      Call ACTUALIZAR_TABLA_MAXREMITO_CUANDO_CAMBIA(FECHAEDITADA%)
    End If
    Call GRACONTROL("", "FENUREFA", FENUE$) 'GUARDO
    
    PREFIREM$ = TRIM$(Text4.TEXT): If PREFIREM$ = "" Then PREFIREM$ = "0001"
    If PREFIREM$ <> "*" Then
      If Len(PREFIREM$) <> 4 Then
        Call COMUNI("Prefijo no Válido")
        Text4.SetFocus
        Exit Sub
      End If
      For U& = 1 To 4
        UI% = Asc(Mid$(PREFIREM$, U&))
        If UI% < 48 Or UI% > 57 Then
            Call COMUNI("Prefijo no Válido")
            Text4.SetFocus
            Exit Sub
        End If
      Next U&
    End If
    Call GRACONTROL("", "PREFIREM", PREFIREM$)
    '
    CMO$ = UCase$(TRIM$(Text5.TEXT))
    Call GRACONTROL("", "CODREFAC", CMO$)
    '
    VERISTOK$ = "NOVERIF"
    If Check1.Value > 0 Then
        VERISTOK$ = "VERISTOK"
    End If
    Call GRACONTROL("REMSTAN", "VERISTOK", VERISTOK$)
    '
    MONUME$ = "NO"
    If Check2.Value = 1 Then
        MONUME$ = "SI"
    End If
    Call GRACONTROL("REMSTAN", "REPLAORI", MONUME$)
    '
'    MONUME$ = "NO"
'    If Option6.Value = True Then
'        MONUME$ = "SI"
'    End If
'    Call GRACONTROL("REMSTAN", "AGRECOOR", MONUME$)

    MONUME$ = "NO"
    If Check3.Value = 1 Then
        MONUME$ = "SI"
    End If
    Call GRACONTROL("REMSTAN", "NUSERIE", MONUME$)
    '
    
    Call GRACONTROL(IDENTER$, "PORTREM", Text12.TEXT)
    Call GRACONTROL("*", "PORTREM", Text12.TEXT)
    
    Call GRACONTROL("", "COPIAREM", Text10.TEXT)
    Call GRACONTROL("REMSTAN", " COCARGA ", Text23.TEXT)
    Call GRACONTROL("REMSTAN", "LIMITEM3", Text11.TEXT)
      
    '
    NROPRO& = 0
    Call NUREMIFA(NROPRO&)
    FOREMIT07.Text11.TEXT = CMO$ + "." + TRIM$(Str$(NROPRO&))
    '
    Call GRACONTROL("*", "PRIETIRE", COMBO1.TEXT)
    Call GRACONTROL(IDENTER$, "PRIETIRE", COMBO1.TEXT)
    
    Call GRACONTROL("", "CONTRODE", TRIM$(Text15.TEXT))
    
    MONUME$ = ""
    If Check4.Value = 1 Then MONUME$ = "SI"
    Call GRACONTROL("CONSLOTE", "SALOTEDE", MONUME$)
    '
    MONUME$ = ""
    If Check5.Value = 1 Then MONUME$ = "SI"
    Call GRACONTROL("REMSTAN", "PRIOBSPE", MONUME$)

    Call GRACONTROL("*", "PRIREEDI", Combo2.TEXT)
    Call GRACONTROL(IDENTER$, "PRIREEDI", Combo2.TEXT)
    '
    Call CIERRARCH("*.*")
    'Hide
    '\\\OT-05-0241-WAR-02/05/05///
    Unload Me 'DESCARGUA EL FORMULARIO DESPUES DEL SETUP
    '\\\OT-05-0241-WAR-FIN///
    '
End Sub

Private Sub Command3_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text1.TEXT = VDEVU$
        Text2.TEXT = VALACT2$("INDIFRM")
    End If
End Sub
'

Private Sub Command4_Click()
    Call SELECHO("SELFECHA")
    FENUE$ = TRIM$(VALACT1$("SELFECHA"))
    If FENUE$ <> "" Then
        Text3.TEXT = FENUE$
    End If
End Sub

Private Sub Command5_Click()
    Call CONECRUN("FLFORMS/" + TRIM$(Text1.TEXT), "Formularios de Impresion")
    If EXISTE%("FLFORMS.EXE") < 1 Then
        Call FINAL("")
    End If
End Sub

Private Sub Command6_Click()
    Call SELECHO("LISTAS")
    VDEVU$ = VALACT1$("LISTAS")
    If VDEVU$ <> "" Then
        Text6.TEXT = VDEVU$
        Text9.TEXT = VALACT2$("LISTAS")
    End If
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   '
   OPX% = ALTERNA%("Recuperacion de Precios Desde Pedidos\Otros Accesos Directos")
   If OPX% < 1 Or OPX% > 2 Then GoTo 99
   
   If OPX% = 1 Then
    If DERACCE%("RECUPREC", "Recuperacion de Precios Desde Pedidos") >= 2 Then
      Screen.MousePointer = 11
      Call FOREMIT07.RECUREMIT
      Screen.MousePointer = 1
      Call COMUNI("Proceso Finalizado")
    End If
   ElseIf OPX% = 2 Then
      Call COMUNI("No Existen Otros Accesos")
   End If
   '
99 Command8.Enabled = True

End Sub

Private Sub Option5_Click()

End Sub

Private Sub Text11_LostFocus()
   Text11 = FORMATNUM$(XVALO(Text11), "F9.2")
End Sub

Private Sub Text13_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text13.TEXT = VDEVU$
        Text14.TEXT = VALACT2$("INDIFRM")
    End If

End Sub

Private Sub Text17_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text17.TEXT = VDEVU$
        Text18.TEXT = VALACT2$("INDIFRM")
    End If

End Sub


Private Sub Text23_LostFocus()
    Text23.TEXT = FORMATNUM$(XVALO(Text23), "F10.2")
End Sub

Private Sub Text7_Change()
        Text8 = ""
        DP1% = XVALO(Text7)
        Text8.TEXT = ECOARCH$("ECODEP", Chr$(DP1%))

End Sub

Private Sub Text9_Click()
   Text6.SetFocus
End Sub

Private Sub Text6_DblClick()
    Call SELECHO("LISTAS")
    VDEVU$ = VALACT1$("LISTAS")
    If VDEVU$ <> "" Then
        Text6.TEXT = VDEVU$
        Text9.TEXT = VALACT2$("LISTAS")
    End If
End Sub

Private Sub Text7_DblClick()
    Call SELECHO("TADEPOSI")
    VDEVU$ = VALACT1$("TADEPOSI")
    If VDEVU$ <> "" Then
        Text7.TEXT = VDEVU$
        Text8.TEXT = VALACT2$("TADEPOSI")
    End If
End Sub
'
Private Sub Command9_Click()
    Call SELECHO("TADEPOSI")
    VDEVU$ = VALACT1$("TADEPOSI")
    If VDEVU$ <> "" Then
        Text7.TEXT = VDEVU$
        Text8.TEXT = VALACT2$("TADEPOSI")
    End If
End Sub
'
Private Sub Form_Load()
    '
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
    
    J& = 0
    For U& = 1 To ULTREG&("SUCURSAL")
        X$ = REGLEIDO$("SUCURSAL", U&)
        CODEPO% = Asc(Mid$(X$, 64, 1))
        If CODEPO% > 0 Then
           If CODEPO% <> 99 Then
              If J& < 62 Then
                 J& = J& + 1
                 Y$ = REGBLAN$("TADEPOSI")
                 Call REPLA(Y$, Chr$(CODEPO%), 1, 1)
                 Call REPLA(Y$, Left$(X$, 63), 2, 63)
                 Call GRAREG("TADEPOSI", Y$, J&)
              End If
           End If
        End If
    Next U&
    Call SETULTREG("TADEPOSI", J&)
    Call CIERRARCH("SUCURSAL")
    Call CIERRARCH("TADEPOSI")
    Screen.MousePointer = 1
    '
    Text1.TEXT = Control$("", "FOREMIFA")
    Text2.TEXT = ECOARCH$("INDIFRM", Text1.TEXT)
    Text3.TEXT = Control$("", "FENUREFA")
    Text7.TEXT = Control$("", "DEPOEXPE")
    Text8.TEXT = ECOARCH$("TADEPOSI", Chr$(Val(Text7.TEXT)))
    Text6.TEXT = Control$("", "LIPREXPE")
    Text9.TEXT = ECOARCH$("LISTAS", Chr$(Val(Text6.TEXT)))
    '
    PREM$ = Control$(IDENTER$, "PORTREM")
      If PREM$ = "" Then PREM$ = Control$("*", "PORTREM")
   Text12.TEXT = PREM$

    Text10.TEXT = Control$("", "COPIAREM")
    '
    Text13.TEXT = Control$("", "FORETIRE")
    Text14.TEXT = ECOARCH$("INDIFRM", Text13.TEXT)
    Text17.TEXT = Control$("", "FORREMSC")
    Text18.TEXT = ECOARCH$("INDIFRM", Text17.TEXT)
    
    Text23.TEXT = Control$("REMSTAN", " COCARGA")
    '
    Text11.TEXT = FORMATNUM$(XVALO(Control$("REMSTAN", " LIMITEM3")), "F9.2")
    If Control$("", "NUREMIFA") = "AUTONUME" Then Option2.Value = True
    If Control$("", "TIPAREMI") = "HOJASUEL" Then Option4.Value = True
    'If CONTROL$("", "MODOEXPE") = "ODESPACH" Then Option6.Value = True
    '
    PREFIREM$ = TRIM$(Control$("", "PREFIREM")): If PREFIREM$ = "" Then PREFIREM$ = "0001"
    Text4.TEXT = PREFIREM$
    CMO$ = TRIM$(Control$("", "CODREFAC"))
    If CMO$ <> "" Then Text5.TEXT = CMO$
    '
    If Control$("REMSTAN", "VERISTOK") = "VERISTOK" Then
       Check1.Value = 1
    End If
    Check7.Value = XVALO(Control$("", "LOTEPARA"))
    '\\\OT-05-0241-WAR-02/05/05///
    'Command19.Enabled = False
    'If Control$("REMSTAN", "REPLAORI") = "SI" Then
    '   Check2.Value = 1
    '   Command19.Enabled = True
    'End If
    If Control$("REMSTAN", "REPLAORI") = "SI" Then 'REEMPLAZA CODIGO POR CODIGO DEL CLIENTE
       Check2.Value = 1
    End If
    '
'    If Control$("REMSTAN", "AGRECOOR") = "SI" Then 'CONCATENA DESCRIPCION CON CODIGO DEL CLIENTE
'       Option6.Value = True
'    End If
    If Control$("REMSTAN", "NUSERIE") = "SI" Then
       Check3.Value = 1
    End If
    
    Check5.Value = 0
    If Control("REMSTAN", "PRIOBSPE") = "SI" Then
      Check5.Value = 1
    End If
    
    Check6.Value = 0 'MOSTRAR DATOS DE ENTREGA DEL PEDIDO SOLO SI HAY UN PEDIDO
    If Control("REMSTAN", "DATENTPE") = "SI" Then
     Check6.Value = 1
    End If
    
    Text15 = Control("", "CONTRODE")
    Text16 = Control("", "PREFSCAR")
    NUMPREF% = XVALO(Text16)
    Command13.Caption = Command13.Caption & " " & NUMPREF%
    
    Dim PRX As Printer
    PREM$ = Control$(IDENTER$, "PRIETIRE")
    If PREM = "" Then PREM$ = Control$("*", "PRIETIRE")
    COMBO1.Clear
    On Error Resume Next
    COMBO1.TEXT = Printer.DeviceName
    On Error GoTo 0
    For Each PRX In Printers
      COMBO1.AddItem PRX.DeviceName
      If PRX.DeviceName = PREM$ Then
        COMBO1.TEXT = PREM$
      End If
    Next
    '
    PREM$ = Control$(IDENTER$, "PRIREEDI")
    If PREM = "" Then PREM$ = Control$("*", "PRIREEDI")
    Combo2.Clear
    On Error Resume Next
    Combo2.TEXT = Printer.DeviceName
    On Error GoTo 0
    For Each PRX In Printers
      Combo2.AddItem PRX.DeviceName
      If PRX.DeviceName = PREM$ Then
        Combo2.TEXT = PREM$
      End If
    Next
    
    Check4.Value = 0
    If Control("CONSLOTE", "SALOTEDE") = "SI" Then
       Command12.Caption = Command12.Caption + " - " + "SI"
       Check4.Value = 1
    End If
    '\\\OT-05-0241-WAR-FIN///
End Sub

Private Sub Text1_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text1.TEXT = VDEVU$
        Text2.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Text2_CLICK()
    Text1.SetFocus
End Sub


Private Sub Text3_DBLCLICK()
    Call SELECHO("SELFECHA")
    FENUE$ = TRIM$(VALACT1$("SELFECHA"))
    If FENUE$ <> "" Then
        Text3.TEXT = FENUE$
    End If
End Sub

Private Sub Text8_Click()
   Text7.SetFocus
End Sub
