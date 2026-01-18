VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form CONPED07 
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta de Pedidos por Número"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   11790
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      Height          =   3975
      Left            =   8160
      TabIndex        =   49
      Top             =   2160
      Width           =   15
   End
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   1440
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   38
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command43 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1420
      Left            =   5360
      Picture         =   "CONPED07.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Actualizar Indice de Selección"
      Top             =   240
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00624E28&
      ForeColor       =   &H80000008&
      Height          =   8415
      Left            =   10830
      ScaleHeight     =   8385
      ScaleWidth      =   1170
      TabIndex        =   29
      Top             =   0
      Width           =   1200
      Begin VB.CommandButton Command10 
         Caption         =   "Print Etiq."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   620
         Left            =   150
         TabIndex        =   58
         Top             =   3920
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   150
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   46
         Top             =   7920
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   47
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   150
         Picture         =   "CONPED07.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Preparación de Pedidos"
         Top             =   6540
         Width           =   650
      End
      Begin VB.CommandButton Command5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   150
         Picture         =   "CONPED07.frx":0A14
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Generar Orden de Fabricación"
         Top             =   5880
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
         Height          =   650
         Left            =   150
         Picture         =   "CONPED07.frx":0D1E
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Despacho - Entrega al Cliente"
         Top             =   7200
         Width           =   650
      End
      Begin VB.CommandButton Command21 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   150
         Picture         =   "CONPED07.frx":1028
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Consultar Pendientes de Despacho"
         Top             =   4560
         Width           =   650
      End
      Begin VB.CommandButton Command23 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   150
         Picture         =   "CONPED07.frx":1332
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Listado Pendientes de Despacho"
         Top             =   5220
         Width           =   650
      End
      Begin VB.CommandButton Command9 
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
         Height          =   650
         Left            =   150
         Picture         =   "CONPED07.frx":163C
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Sale de la Aplicación"
         Top             =   200
         Width           =   650
      End
      Begin VB.CommandButton Command26 
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
         Height          =   650
         Left            =   150
         Picture         =   "CONPED07.frx":1786
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Configurar Comprobante de Despacho"
         Top             =   2160
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command6 
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
         Height          =   650
         Left            =   150
         Picture         =   "CONPED07.frx":1BC8
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Consultas y Re-Impresión"
         Top             =   2840
         Width           =   650
      End
      Begin VB.CommandButton Command1 
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
         Height          =   650
         Left            =   150
         Picture         =   "CONPED07.frx":1ED2
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Graba las Modificaciones"
         Top             =   860
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command2 
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
         Height          =   650
         Left            =   150
         Picture         =   "CONPED07.frx":21DC
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   1520
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command8 
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
         Height          =   330
         Left            =   150
         TabIndex        =   30
         Top             =   3550
         Visible         =   0   'False
         Width           =   650
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E2D3C0&
      Caption         =   "PEDIDO / ORDEN DE VENTA"
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
      Height          =   1550
      Left            =   105
      TabIndex        =   10
      Top             =   120
      Width           =   5250
      Begin VB.CommandButton Command11 
         Caption         =   "!"
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
         Left            =   2470
         TabIndex        =   36
         ToolTipText     =   "Pedidos por Vendedor"
         Top             =   315
         Width           =   300
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
         Height          =   300
         Left            =   2130
         TabIndex        =   13
         Top             =   315
         Width           =   300
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Height          =   285
         Left            =   4830
         TabIndex        =   25
         ToolTipText     =   "Doble Click para Restaurar Situación de Preparacion y Despacho"
         Top             =   315
         Width           =   330
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "St"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4530
         TabIndex        =   24
         Top             =   360
         Width           =   210
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Height          =   285
         Left            =   3480
         TabIndex        =   23
         Top             =   1125
         Width           =   1680
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Orden Coimpra"
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
         Left            =   2835
         TabIndex        =   22
         Top             =   1020
         Width           =   660
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Height          =   285
         Left            =   1155
         TabIndex        =   19
         Top             =   1125
         Width           =   1605
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   315
         TabIndex        =   18
         Top             =   800
         Width           =   750
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Height          =   285
         Left            =   1155
         TabIndex        =   17
         Top             =   720
         Width           =   4020
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2640
         TabIndex        =   16
         Top             =   420
         Width           =   675
      End
      Begin VB.Label Label7 
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
         Height          =   300
         Left            =   1155
         TabIndex        =   15
         Top             =   315
         Width           =   940
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Height          =   285
         Left            =   3360
         TabIndex        =   14
         Top             =   315
         Width           =   1065
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Venta / Reparación"
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
         Index           =   1
         Left            =   -240
         TabIndex        =   12
         Top             =   1000
         Width           =   1305
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Número"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   -840
         TabIndex        =   11
         Top             =   420
         Width           =   1905
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E2D3C0&
      Caption         =   "ENTREGA"
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
      Height          =   1550
      Left            =   5725
      TabIndex        =   7
      Top             =   120
      Width           =   5000
      Begin VB.Label LENTREGA 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Height          =   285
         Index           =   2
         Left            =   1155
         TabIndex        =   28
         Top             =   1125
         Width           =   3705
      End
      Begin VB.Label LENTREGA 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Height          =   285
         Index           =   1
         Left            =   1155
         TabIndex        =   27
         Top             =   720
         Width           =   3705
      End
      Begin VB.Label LENTREGA 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Height          =   285
         Index           =   0
         Left            =   1155
         TabIndex        =   26
         Top             =   315
         Width           =   3705
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Domicilio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   -840
         TabIndex        =   9
         Top             =   420
         Width           =   1905
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Transporte"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   5
         Left            =   -840
         TabIndex        =   8
         Top             =   1230
         Width           =   1905
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2D3C0&
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
      Left            =   105
      TabIndex        =   2
      Top             =   6300
      Width           =   10620
      Begin VB.Frame Frame8 
         BackColor       =   &H00E2D3C0&
         Height          =   495
         Left            =   0
         TabIndex        =   54
         Top             =   230
         Width           =   10600
         Begin VB.Label Label5 
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
            Left            =   8400
            TabIndex        =   57
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label4 
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
            Left            =   1200
            TabIndex        =   56
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
            Left            =   120
            TabIndex        =   55
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.TextBox Text1 
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
         TabIndex        =   3
         Top             =   735
         Width           =   7155
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
         Height          =   745
         Left            =   8360
         Style           =   1  'Graphical
         TabIndex        =   5
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
         TabIndex        =   45
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   4
         Top             =   720
         Width           =   1170
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2D3C0&
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
      Left            =   105
      TabIndex        =   1
      Top             =   1800
      Width           =   10620
      Begin VB.Frame Frame3 
         Height          =   3975
         Left            =   1120
         TabIndex        =   48
         Top             =   360
         Width           =   15
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00E2D3C0&
         Height          =   495
         Left            =   0
         TabIndex        =   50
         Top             =   240
         Width           =   10600
         Begin VB.Label Label10 
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
            Left            =   120
            TabIndex        =   53
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label21 
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
            Left            =   1200
            TabIndex        =   52
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label22 
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
            Left            =   8160
            TabIndex        =   51
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
         Height          =   3630
         Left            =   15
         TabIndex        =   6
         Top             =   735
         Width           =   10575
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "CONPED07.frx":24E6
      Top             =   1800
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   300
      OleObjectBlob   =   "CONPED07.frx":271A
      TabIndex        =   39
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion:"
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
      Left            =   -105
      TabIndex        =   0
      Top             =   2625
      Visible         =   0   'False
      Width           =   1695
   End
End
Attribute VB_Name = "CONPED07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command10_Click()

  Command10.Enabled = False
  ETIQAHP.Show 1
  GoTo 99
  
  Call CARGA_ENCABEZADO(FRMZELECHO.msf2, "Cuenta/A10;Razon Social/a48", FRMZELECHO)
  Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "Cuenta/A10;Razon Social/a48", FRMZELECHO)
  FRMZELECHO.Caption = "Clientes Activos"

  Dim obj As New RECORXET
  Set rs1 = obj.RS_DEUDOR12()
  Call Carga_MSF_Recordset(rs1, "Cuenta/A10;Razon Social/a48", FRMZELECHO.msf2)
  Set obj = Nothing
  rs1.Close
  Set rs1 = Nothing

  FRMZELECHO.Width = 8500
  FRMZELECHO.Show 1
  
  NCLI = XVALO(RESP_ZELE_VECT(1))
  If NCLI < 1 Then GoTo 99



    Campos$ = "Cta/F6;R.Social/A24;F.Ped./D8;Nro.Ped./F7;Código/A16;Descripición/A24;Cantidad/f12;F.Sol./D8;Ord.Compra/A16"
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO)
    Dim obj2 As New RECORXET
    Set rs1 = obj2.RS_PEDIPEN2(0, NCLI) ' FUNCION QUE DEVUELVE UN RECORDSET
    i& = 0
    With rs1
      Do While Not .EOF
          CIXI% = CODINT%(SAFETEXT$(!CODIEXT))
          If CIXI% < 1 Then GoTo 30
'          If APROBADO1%(CIXI%) < 1 Then GoTo 30 ' NEXT
          NC = XVALO(!NROCLIE)
          RASO$ = rasocli$(NC)
          NPED& = XVALO(!NROPED)
          FEMI% = CVINT(!FECHEMIS)
          DESC$ = DESCRIT0$(CIXI%)
          CantSol = XVALO(!CanPed)
          FentSol% = CVINT(!FeSolEnt)
          i& = i& + 1
          FRMZELECHO.msf2.Rows = i& + 1
          FRMZELECHO.msf2.TextMatrix(i&, 1) = XVALO(!NROCLIE)
          FRMZELECHO.msf2.TextMatrix(i&, 2) = SAFETEXT$(!RaSoClie)
          FRMZELECHO.msf2.TextMatrix(i&, 3) = FECHATEX$(FEMI%)
          FRMZELECHO.msf2.TextMatrix(i&, 4) = NPED&
          FRMZELECHO.msf2.TextMatrix(i&, 5) = CODEXT$(CIXI%)
          FRMZELECHO.msf2.TextMatrix(i&, 6) = DESC$
          FRMZELECHO.msf2.TextMatrix(i&, 7) = FORMATNUM$(CantSol, "f10.1")
          FRMZELECHO.msf2.TextMatrix(i&, 8) = FECHATEX$(FentSol%)
          FRMZELECHO.msf2.TextMatrix(i&, 9) = OrdenCompraPed$(NPED&)
30       .MoveNext
      Loop
    End With
    On Error Resume Next
    Set obj2 = Nothing
    rs1.Close
    Set rs1 = Nothing
    On Error GoTo 0
    
    FRMZELECHO.Width = 14500
    FRMZELECHO.Caption = "Items C/Pedidos Pendientes Sin Aprobar - (Doble Click Sobre Items Para Ver Ficha Técnica y Aprobar)"
    Screen.MousePointer = 1
    FRMZELECHO.Show 1
    '
    CIXI% = CODINT%(SAFETEXT$(RESP_ZELE_VECT(5)))
    If CIXI% < 1 Then Exit Sub
    CODI$ = SAFETEXT$(RESP_ZELE_VECT(5))
    
    FRMENVAP.TXTCODIGO = CODI$
    FRMENVAP.Show 1
    Unload VENTABNEW
    Dim strArray() As String

    Campos$ = "FECHA/D8;ID_MATER./F8;MATERIA PRIMA/A18;DATOS/A18;DESCRIPCIÓN/A32;N° BOBINAS - LOTE/A18;PROV./F10;R.Social/A10;ANCHO/F12.2;GRJE - MC./F12.2;COLOR/A12;MTS./f12.2;CONSUMO/f12.2;SALDO/f12.2;Obsesrvaciones/A48;ULT/A0"
    VENTANA.Inicializar = 0
    VENTANA.Campox = Campos$
    VENTANA.AgregaRegistro = 1
    VENTANA.UtilizaFormula = 1
    'VENTANA.formula(15) = "=EXCEL;(L1/K1)*(L1/K1)*M1"
    'VENTANA.formula(14) = "=EXCEL;(12/11)*(12/11)*13"
    VENTANA.Formula(14) = "12;-;13"
    VENTANA.HabilitaBorrar = 1
    VENTANA.HabilitaInsertar = 0
    VENTANA.NoMostrarBotonTilde = 0
    VENTANA.TITULO = "CONSUMO"
    
    VENTANA.ColEcoCampo(2) = "CAMPO1;CAMPO2;DATASQL;LISTA = 'MATERIALES ENVAPLAST'"
    VENTANA.ColEcoCampo(4) = "CODIGO;DESCRIPCION;MASTER"
    VENTANA.Condicion2(4) = "FAMILIA = COL@2" ' CONDICION FILTRADA POR COLUMNA 1

    VENTANA.ColEcoCampo(7) = "NUME_CLI;Fant_Cli;Proved12"
    'VENTANA.CampEditables = "1;2;4;5;6;7;8;9;10;11;12"
    VENTANA.CampEditables = "1;2;4;6;7;12;13;15"
    VENTANA.Inicializar = 1
    '
'    Call CopiarMsfAOtroMsf(MSF12(Index), VENTABNEW.MSF, 1)
'    '****************************************************************
'    Call T_Espera(1 / 10) ' PARA QUE EL DOBLE CLICK NO APLIQUE SOBRE LA VENTANA QUE SE CARGA
'30  VENTABNEW.Show 1
'    If VENTABNEW.APROBADO% > 0 Then
'       MODIFICADO(Index) = 1
'       If FilaVacia%(VENTABNEW.MSF) < 1 Then
'        If FECHANUM(VENTABNEW.MSF.TextMatrix(1, 1)) < 1 Then
'           Call COMUNI("Fecha Inválida")
'           GoTo 30
'        End If
'       End If
'       Call CopiarMsfAOtroMsf(VENTABNEW.MSF, MSF12(Index), 1)
'       Me.TXTCONSUMO3(Index) = FORMATNUM$(XVALO(SUMA_COLUMNA_SELECCION(MSF12(GLOBALFRM_SOLASUP%), 13)), "F12.2")
'       For i& = 1 To VENTABNEW.MSF.Rows - 1
'          If TRIM$(VENTABNEW.MSF.TextMatrix(i&, 5)) <> "" Then 'campo de la descripcion del codigo
'            CODIX$ = TRIM$(VENTABNEW.MSF.TextMatrix(i&, 4))
'            If CODIX_ANT$ <> CODIX$ Then
'              CI1% = CODINT(CODIX$)
'              strArray = Split(AnchoColorGramaje$(CI1%), ";")
'            End If
'            MSF12(Index).TextMatrix(i&, 9) = TRIM$(FORMATNUM$(XVALO(strArray(0)), "F16.4"))
'            MSF12(Index).TextMatrix(i&, 10) = TRIM$(FORMATNUM$(XVALO(strArray(1)), "F16.4"))
'            MSF12(Index).TextMatrix(i&, 11) = TRIM$(strArray(2))
'            CODIX_ANT$ = TRIM$(VENTABNEW.MSF.TextMatrix(i&, 4))
'          End If
'       Next i&
'
'    End If
'    Erase strArray
'    Set VENTANA = Nothing
'    Unload VENTABNEW
99 Command10.Enabled = True
End Sub




Private Sub Command11_Click()
  Command11.Enabled = False
  '
  Call SELECHO("VENDEDOR")
  CodVend$ = VALACT1$("VENDEDOR")
  DESVEND$ = VALACT2$("VENDEDOR")
  If CodVend$ <> "" Then
     '
     'Se arma la lista de pedidos en base al vendedor seleccionado
     CONDI$ = "NROPED > 0 AND VENDED = " & CodVend$ & " ORDER BY NroPed DESC"
     Call CARGALISEL("!INDIPEDI", "PEDENCA", "NROPED/F6.0 ; REFERENCIA/A60", CONDI$)
     '
     Call SELECHO("!INDIPEDI/Indice General de Pedidos (Pend.-Cump.-Anul.) - " & DESVEND$)
     CODPED$ = VALACT1$("!INDIPEDI")
     '
     If CODPED$ <> "" Then
        Label7 = CODPED$
     End If
  End If
  '
  Command11.Enabled = True

End Sub

Private Sub Command21_Click()
   Command21.Enabled = False
   Call COPYSTRU("INDIPEDI", "PEPENDES")
   '
   Call PoneEnCeroNull("peddeta", "AcuDespa", REGAF&)
   Call PoneEnCeroNull("peddeta", "VenRep", REGAF&)
   Call PoneEnCeroNull("peddeta", "CodiEmpr", REGAF&)
   Call PoneEnCeroNull("peddeta", "StaDespa", REGAF&)
   Call PoneEnCeroNull("PEDENCA", "CodiEmpr", REGAF&)
   '
   Set IPEPE = New ADODB.Recordset
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If InStr(EPAC$, "AMBEST") > 0 Then
        SQLX$ = "SELECT DISTINCT PEDDETA.NROPED, PEDENCA.REFERENCIA FROM PEDDETA WITH(NOLOCK)"
        SQLX$ = SQLX$ + " INNER JOIN PEDENCA"
        SQLX$ = SQLX$ + " ON PEDDETA.NroPed = PEDENCA.NroPed"
        SQLX$ = SQLX$ + " WHERE PEDDETA.NroPed > 0 "
        SQLX$ = SQLX$ + " AND PEDENCA.Status < 9 "                ' NO ANULADOS
        SQLX$ = SQLX$ + " AND PEDDETA.Status <> 8 "
        SQLX$ = SQLX$ + " AND PEDDETA.Status <> 9 "
        SQLX$ = SQLX$ + " AND PEDENCA.CODIEMPR = " & CodiEmp% & " "
        SQLX$ = SQLX$ + " AND PEDDETA.CODIEMPR = " & CodiEmp% & " "
        SQLX$ = SQLX$ + " AND PEDDETA.CanPed > PEDDETA.AcuDespa"
        SQLX$ = SQLX$ + " AND PEDDETA.CANPED > PEDDETA.CANTENT"
        SQLX$ = SQLX$ + " AND PEDDETA.VenRep = 'V' "             ' Excluye Reparaciones
        SQLX$ = SQLX$ + " ORDER BY PEDDETA.NroPed ASC "
        GoTo 25
   End If

   SQLX$ = "SELECT DISTINCT PEDDETA.NROPED, PEDENCA.REFERENCIA FROM PEDDETA WITH(NOLOCK)"
   SQLX$ = SQLX$ + " INNER JOIN PEDENCA"
   SQLX$ = SQLX$ + " ON PEDDETA.NroPed = PEDENCA.NroPed"
   SQLX$ = SQLX$ + " WHERE PEDDETA.NroPed > 0 "
   SQLX$ = SQLX$ + " AND PEDENCA.Status < 9 "                ' NO ANULADOS
   SQLX$ = SQLX$ + " AND PEDDETA.Status <> 8 "
   SQLX$ = SQLX$ + " AND PEDDETA.Status <> 9 "
   SQLX$ = SQLX$ + " AND PEDDETA.StaDespa < 2 " ' elimina los cancelados
   SQLX$ = SQLX$ + " AND PEDENCA.CODIEMPR = " & CodiEmp% & " "
   SQLX$ = SQLX$ + " AND PEDDETA.CODIEMPR = " & CodiEmp% & " "
   SQLX$ = SQLX$ + " AND PEDDETA.CanPed > PEDDETA.AcuDespa"
   SQLX$ = SQLX$ + " AND PEDDETA.VenRep = 'V' "             ' Excluye Reparaciones
   SQLX$ = SQLX$ + " ORDER BY PEDDETA.NroPed ASC "
   
'25 On Error Resume Next
25 Set IPEPE = READSET(SQLX$)

''   IPEPE.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
'   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'      On Error GoTo 0
'      Call CapturaErrorOpen
'      GoTo 25
'   End If
'   On Error GoTo 0
   '
   JJ& = 0
   JK& = 0: NPEDXA& = 0: KU& = 0
   REBLA$ = REGBLAN$("CONPEDIP")
   With IPEPE
'     On Error Resume Next OD-30/11/11 MODIFICADO POR QUE DABA ERROR EN EL MOVEFIRST
'     .MoveLast
'     If Not (Err) Then
'        On Error GoTo 0
'        FIN& = .RecordCount
'        .MoveFirst
        Do While Not .EOF
           NPX& = !NROPED
           DPX$ = !REFERENCIA
           YY$ = MKS$(NPX&) + DPX$
           JJ& = JJ& + 1
           Call GRAREG("PEPENDES", YY$, JJ&)
        .MoveNext
        Loop
'     End If
   End With
   '
   Call FILESORT("PEPENDES", "", 1, 1, "ASC")
   Call SETULTREG("PEPENDES", JJ&)
   Call FRESHECHO("PEPENDES")
   If JJ& < 1 Then
     Call COMUNI("No hay Pedidos\Pendientes de Despacho.")
     GoTo 99
   End If
   '
3  Call SELECHO("PEPENDES/Indice de Pedidos Pendientes de Despacho)")
   NPX& = Val(VALACT1$("PEPENDES"))
   If NPX& < 1 Then GoTo 99
   '
   Label7 = TRIM$(Str$(NPX&))
   '
99 Command21.Enabled = True
End Sub

Private Sub Command23_Click()
   Command23.Enabled = False
   Dim CIKU%(4196), NPKU&(4196), CAYADES(4196)
   '
   OPLIDE07.Show 1
   Call FRESHALL
   DoEvents
   If OPLIDE07.OPIMPRE = "" Then GoTo 99
   '
   Screen.MousePointer = 11
   NEPREPA% = 0
   If OPLIDE07.Check1(1).Value = 1 Then NEPREPA% = 1
   '
   Dim ACTI%(32767), STACT(32767), OFPEN(32767)
   Dim NROPEPE&(8192), REGPEPE&(8192), SALPEPE(8192)
   Dim ITPEDI$(), CODIT%(), CAPREPA(), CAYAEN(), NUPEDI&()
   MAXPED& = 4096
   ReDim ITPEDI$(MAXPED&), CODIT%(MAXPED&), CAPREPA(MAXPED&), CAYAEN(MAXPED&), NUPEDI&(MAXPED&)
   '
   DEPEX% = Int(Val(Control$("", "DEPOEXPE")))
   DETER% = Int(Val(Control$("", "DEPOTERM")))
   If DEPEX% < 1 Or DETER% < 1 Or DEPEX% = DETER% Then
     Call MENSERR(24, "Proceso Imposible.\  \Falta Configurar Depósitos\de Expedición\y Productos Terminados.")
     GoTo 99
   End If
   '
'   SQLX$ = "SELECT PEDDETA.FESOLENT AS FESOLE, * FROM PEDDETA WITH(NOLOCK)"
'   SQLX$ = SQLX$ + " INNER JOIN PEDENCA"
'   SQLX$ = SQLX$ + " ON PEDDETA.NroPed = PEDENCA.NroPed"
'   SQLX$ = SQLX$ + " WHERE PEDDETA.NroPed > 0 "
'   SQLX$ = SQLX$ + " AND PEDENCA.Status < 9 "                ' NO ANULADOS
'   SQLX$ = SQLX$ + " AND PEDDETA.Status <> 8 "
'   SQLX$ = SQLX$ + " AND PEDDETA.Status <> 9 "
'   SQLX$ = SQLX$ + " AND PEDDETA.StaDespa < 2 "              ' elimina los cancelados
'   SQLX$ = SQLX$ + " AND PEDDETA.CanPed > PEDDETA.AcuDespa"
'   SQLX$ = SQLX$ + " AND PEDDETA.VenRep = 'V' "             ' Excluye Reparaciones
'   SQLX$ = SQLX$ + " ORDER BY PEDDETA.NroPed ASC, PEDDETA.NuOrItem ASC"
'   Set IPEPE = New ADODB.Recordset
'   IPEPE.Open (SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
   CONDI$ = " NroPed > 0 "
   CONDI$ = CONDI$ + " AND Status < 8 "
   CONDI$ = CONDI$ + " AND StaDespa < 2 "
   CONDI$ = CONDI$ + " AND CanPed > AcuDespa"
   CONDI$ = CONDI$ + " AND VenRep = 'V' "             ' Excluye Reparaciones
   
   SQLX$ = "SELECT NROPED, Status, VenRep,NroClie,CodiInt,CanPed,FechEmis, FeSolEnt, AcuPrepa, AcuDespa, NROOCOM "
   SQLX$ = SQLX$ + ",TColor , TLargo, TAncho, TEspesor, TTalle  FROM PEDDETA WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE  " & CONDI$
'''   SQLX$ = SQLX$ + " AND PEDENCA.Status < 9 "                ' NO ANULADOS
''   SQLX$ = SQLX$ + " AND Status < 8 "
''   SQLX$ = SQLX$ + " AND StaDespa < 2 "              ' elimina los cancelados
''   SQLX$ = SQLX$ + " AND CanPed > AcuDespa"
''   SQLX$ = SQLX$ + " AND VenRep = 'V' "             ' Excluye Reparaciones
   SQLX$ = SQLX$ + " ORDER BY NroPed ASC, NuOrItem ASC"
   '
   FIN& = CantRegistros("PEDDETA", CONDI$, "NOMESS")
   If FIN& < 1 Then
           Call MENSERR(24, "Aparente Error de Acumulación y Despacho.\   \Para Recuperar Ingrese a Preparación de Pedidos.\   \Luego Vuelva a Pedir Listado.")
           Screen.MousePointer = 1
           GoTo 99
   End If

   Set IPEPE = READSET(SQLX$)
'   Dim IPEPE As ADODB.Recordset
'   Set IPEPE = New ADODB.Recordset
'25 On Error Resume Next
'   IPEPE.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
'   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'      On Error GoTo 0
'      Call CapturaErrorOpen
'      GoTo 25
'   End If
'   On Error GoTo 0
   '
   JJ& = 0
   JK& = 0: NPEDXA& = 0: KU& = 0
   REBLA$ = REGBLAN$("CONPEDIP")
   With IPEPE
'     On Error Resume Next
'     .MoveLast
'     If Not (Err) Then
'        On Error GoTo 0
'        FIN& = .RecordCount
'        If FIN& < 1 Then
'           Call MENSERR(24, "Aparente Error de Acumulación y Despacho.\   \Para Recuperar Ingrese a Preparación de Pedidos.\   \Luego Vuelva a Pedir Listado.")
'           Screen.MousePointer = 1
'           GoTo 99
'        End If
        ULTIPE& = XVALO(!NROPED)
        Call COMUNI("La Ultima O/Venta (Venta) Registrada\es la Numero " + TRIM$(Str$(ULTIPE&)))
        '
        .MoveFirst
        Do While Not .EOF
        KU& = KU& + 1
        Call TRACE(20, KU&, FIN&)
        NPEDXX& = XVALO(!NROPED)
        STATX% = XVALO(!Status)
        REPARX% = 0: If SAFETEXT$(!VenRep) = "R" Then REPARX% = 1
        '
        If REPARX% <> 1 Then
          STAENCA% = XVALO(VALOBADA("PEDENCA", "STATUS", "NroPed = " & NPEDXX&))
          If STAENCA% < 9 Then
             NCLIE = XVALO(!NROCLIE)
             CIXI% = XVALO(!CODIINT)
             CAPEDID = XVALO(!CanPed)                  '  CANTIDAD PEDIDA
             FEMIPED% = CVINT(!FECHEMIS)
             FENTRESOL% = CVINT(!FeSolEnt)
             If CIXI% > 0 And CIXI% <= NUMAS% Then
                CAPRE = XVALO(!AcuPrepa)
                CAYEN = XVALO(!AcuDespa)
                '
                ' ESTO ES PARA COMPENSAR LA APARICION DE PEDIDOS VIEJOS - EPB 28/02/06
                If CAPRE < CAYEN Then
                   CAPRE = CAYEN
                End If
                '
                CASAL = CAPEDID - CAYEN
26              If NEPREPA% > 0 Then CASAL = CAPEDID - CAPRE
                CAPEN = CASAL
                '
                If CASAL > 0.05 Then
                   YY$ = REBLA$
                   Call REPLA(YY$, MKS$(!NROPED), 1, 4)
                   Call REPLA(YY$, MKI$(FEMIPED%), 5, 2)
                   Call REPLA(YY$, MKI$(NCLIE), 7, 2)
                   Call REPLA(YY$, !NROOCOM, 9, 16)
                   Call REPLA(YY$, MKI$(FENTRESOL%), 25, 2)
                   Call REPLA(YY$, MKI$(!CODIINT), 33, 2)
                   Call REPLA(YY$, !TCOLOR, 35, 16)
                   Call REPLA(YY$, !TTALLE, 51, 8)
                   Call REPLA(YY$, MKS$(!TLargo), 59, 4)
                   Call REPLA(YY$, MKS$(!TAncho), 63, 4)
                   Call REPLA(YY$, MKS$(!TEspesor), 67, 4)
                   Call REPLA(YY$, MKS$(CASAL), 71, 4)
                   Call REPLA(YY$, String$(54, 0), 75, 54)
                   NPEDX& = CVS(Left$(YY$, 4))
                   STAPEA% = 1
                   If NPEDX& <> NPEDXA& Then STAPEA% = 1
                   '
                   If ACTI%(CIXI%) < 1 Then
                     ACTI%(CIXI%) = 1
                     STACT(CIXI%) = STOCKACT(CIXI%)
                     '
                     If NEPREPA% = 1 Then
                       If DEPEX% > 0 Then STACT(CIXI%) = STACT(CIXI%) - STODEPOS(CIXI%, DEPEX%)
                     End If
                     '
                     If STACT(CIXI%) < 0 Then STACT(CIXI%) = 0
                     OFPEN(CIXI%) = STORFAB(CIXI%)
                   End If
                   '
                   Call REPLA(YY$, MKS$(STACT(CIXI%)), 75, 4)
                   Call REPLA(YY$, MKS$(OFPEN(CIXI%)), 79, 4)
                   FALTA = CAPEN - STACT(CIXI%) - OFPEN(CIXI%)
                   If FALTA < 0 Then FALTA = 0
                   Call REPLA(YY$, MKS$(FALTA), 83, 4)
                   JJ& = JJ& + 1
                   Call GRAREG("CONPEDIP", YY$, JJ&)
                   '
                   STAPE% = 1
                   If CAPEN > CVS(Mid$(YY$, 75, 4)) Then STAPE% = 2
                   If FALTA > 0 Then STAPE% = 3
                   If STAPE% > STAPEA% Then STAPEA% = STAPE%
                   '
                   For KK& = JJ& To 1 Step -1
                     ZZ$ = REGLEIDO$("CONPEDIP", KK&)
                     If CVS(Left$(ZZ$, 4)) <> NPEDX& Then GoTo 15
                     Call REPLA(ZZ$, MKI$(STAPEA%), 127, 2)
                     Call GRAREG("CONPEDIP", ZZ$, KK&)
                   Next KK&
                   '
15                 NPEDXA& = NPEDX&
                   If CAPEN <= STACT(CIXI%) Then
                       STACT(CIXI%) = STACT(CIXI%) - CAPEN
                       CAPEN = 0
                     Else
                       CAPEN = CAPEN - STACT(CIXI%)
                       STACT(CIXI%) = 0
                   End If
                   '
                   If CAPEN <= OFPEN(CIXI%) Then
                       OFPEN(CIXI%) = OFPEN(CIXI%) - CAPEN
                       CAPEN = 0
                     Else
                       CAPEN = CAPEN - OFPEN(CIXI%)
                       OFPEN(CIXI%) = 0
                   End If
                   '
                End If
             End If
          End If
        End If
        .MoveNext
        Loop
'     End If
   End With
   '
   Call SETULTREG("CONPEDIP", JJ&)
   Screen.MousePointer = 1
   If JJ& < 1 Then
     Call COMUNI("No hay Pedidos\Pendientes de Despacho.")
     GoTo 99
   End If
   '
   Call HEADERS("CONPEDIP", "")
   MAGRU% = Val(OPLIDE07.AGRUPA)
   If MAGRU% = 1 Then
       OPX% = Val(OPLIDE07.OPIMPRE)
       If OPX% > 3 Then GoTo 99
       '
       If OPX% < 1 Then
            SITU$ = "Todos los Pendientes": If NEPREPA% = 1 Then SITU$ = SITU$ + " - Netos de Preparados"
            Call HEADERS("CONPEDIP", "Situación: " + SITU$)
          Else
            KK& = 0
            For U& = 1 To JJ&
              YY$ = REGLEIDO$("CONPEDIP", U&)
              If CVI(Mid$(YY$, 127, 2)) = OPX% Then
                KK& = KK& + 1
                Call GRAREG("CONPEDIP", YY$, KK&)
              End If
            Next U&
            Call SETULTREG("CONPEDIP", KK&)
            Call CIERRARCH("CONPEDIP")
            If OPX% = 1 Then
                 SITU$ = "Listos para Despachar": If NEPREPA% = 1 Then SITU$ = SITU$ + " - Netos de Preparados"
                 Call HEADERS("CONPEDIP", "Situación: " + SITU$)
              ElseIf OPX% = 2 Then
                 SITU$ = "En Preparación / Fabricación": If NEPREPA% = 1 Then SITU$ = SITU$ + " - Netos de Preparados"
                 Call HEADERS("CONPEDIP", "Situación: " + SITU$)
              ElseIf OPX% = 3 Then
                 SITU$ = "Para Fabricar": If NEPREPA% = 1 Then SITU$ = SITU$ + " - Netos de Preparados"
                 Call HEADERS("CONPEDIP", "Situación: " + SITU$)
            End If
       End If
       Call FINAL("*LIDESPE")
     ElseIf MAGRU% = 2 Then
       Call FILESORT("CONPEDIP", "", 7, 7, "ASC")
       Call FINAL("*LIDESQU")
   End If
   '
99 Command23.Enabled = True
   '
End Sub

Private Sub Command24_Click()
   Command24.Enabled = False
   If DERACCE%("DESPAMAT", "Despachar Pedidos de Clientes") > 1 Then
     Call DESPAMAT
     Call Label7_Change
   End If
   Command24.Enabled = True
End Sub

Private Sub Command26_Click()
   Command26.Enabled = False
   If DERACCE%("STPSEPED", "Configuración de Seguimiento de Pedidos") = 2 Then
      CFGDES07.Show 1
   End If
   Command26.Enabled = True
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   NPX& = Val(Label7)
   If NPX& > 0 Then
     If TRIM$(Text1) <> "" Then
       Call FRATEXTO(Text1, 56)
       
       For KKI% = 1 To CARETEX%
         TTX1$ = Space$(76)
         Call REPLA(TTX1$, RETEX$(KKI%), 11, 56)
         If KKI% = 1 Then
           Call REPLA(TTX1$, HOS$, 1, 8)
           Call REPLA(TTX1$, USERTER$, 69, 24)
         End If
         'Call REGAPP("ANOTAPED", Left$(MKS$(NPX&) + TTXX$, 96))
         TTXX$ = TTXX$ + TTX1$ + Chr$(10) + Chr$(13)
       Next KKI%
       '
       '\\\OT-08-0098-12/03/08-///
       SQLX$ = "SELECT ANOTAPED FROM PEDENCA"
       SQLX$ = SQLX$ + " WHERE NroPed = " & NPX& & ""
       Dim PEDDDENCA As New ADODB.Recordset
       Set PEDDDENCA = New ADODB.Recordset
25     On Error Resume Next
       PEDDDENCA.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 25
       End If
       On Error GoTo 0
       '
       With PEDDDENCA
          If Not (.BOF And .EOF) Then
             ANOTA$ = !AnotaPed
             ANOTA$ = ANOTA$ + TTXX$
             ANOTA$ = ANOTA$ + String$(88, "=") + Chr$(10) + Chr$(13)
             !AnotaPed = ANOTA$
             .Update
           End If
       End With
       PEDDDENCA.Close
       Set PEDDDENCA = Nothing
       '
       Text1 = ""
       '
       List1.Clear
       Label7 = "": Label7 = NPX&
       '
     End If
   End If
   Command3.Enabled = True
End Sub

Private Sub Command4_Click()
   Command1.Enabled = False
   '
   Call CARGAPEDIDOS(1, 1)
   NPX& = XVALO(RESP_ZELE_VECT(1))
   Label7 = TRIM$(Str$(NPX&))
'   ' Call CARGALISEL("INDIPEDI", "PEDENCA", "NROPED/F6.0 ; REFERENCIA/A60", "NROPED > 0 ORDER BY NROPED DESC")
'3  Call SELECHO("INDIPEDI/Indice General de Pedidos (Pendientes-Cumplidos-Anulados)")
'   NPX& = Val(VALACT1$("INDIPEDI"))
'   If NPX& < 1 Then GoTo 99
'   '
'   Label7 = TRIM$(Str$(NPX&))
   '
99 Command1.Enabled = True
End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
   ACONE$ = "ORFABR07"
   Call CONECRUN(ACONE$, "Ordenes de Fabricación")
   Command5.Enabled = True
End Sub

Sub Command6_Click()
      Command6.Enabled = False
      NPX& = Val(Label7)
      '
      If NPX& < 1 Then GoTo 99
      NUDOCU$ = TRIM$(Str$(NPX&))
      '
10    OPX% = ALTERNA%("Visualizar Pedidos - Orden de Venta\Visualizar Orden de Despacho\Detalle de B.Datos\Estado de Pedido\Refrescar Pedidos")
      If OPX% < 1 Or OPX% > 5 Then GoTo 99
      '
      If OPX% = 1 Or OPX% = 3 Then
        If DERACCE%("VERCOSTO/NOMESS", "Visualizar Costos y Precios") < 1 Then
           Call MENSERR(24, "Usuario No Autorizado para Visualizar Pedidos.\Seleccione Orden de Despacho")
           GoTo 10
        End If
      End If
      '
      If OPX% = 3 Then
        ABUS$ = NUDOCU$
        Call CONECRUN("GENREP07/DETAPEDNRO", "Consulta por Número de Remito")
        GoTo 99
      End If
      '
      If OPX% = 1 Then
         FORIPRE$ = TRIM$(Control$("", "FORPECLI"))
         ARCSQR$ = FORIPRE$
         If EXISTE%(LUDAT$ + ARCSQR$ + ".SQR") > 0 Then
            NPSEL& = XVALO(Label7)
            NCLIEN = cliente_SEGUN_PEDIDO(NPSEL&)
            FILTX$ = TRIM$(Str$(NPSEL&)) & ";" & TRIM$(Str$(NCLIEN))
            Call STDFORM(ARCSQR$, FILTX$)
            GoTo 99
         End If
      
         TIDOCUM$ = "Pedido de Cliente"
         DOCUMENTO_ELEGIDO$ = "Orden de Venta"
         'BUSCA EL DOCUMENTO EN LA TABLA SQL CON LA NUEVA FUNCION
25       OPX% = COMALTER%("Opciones de Visualización de Pedidos\\Ver Última Impresión\Seleccionar Versión")
         If OPX% = 2 Then
30            DOCUSEL& = ListaSeleccionDocumVersionados&(XVALO(NUDOCU$), TIDOCUM$)
            If DOCUSEL& > 0 Then
               Call MUESTRADOC(DOCUSEL&, "")
               GoTo 30
            Else
               GoTo 25
            End If
         ElseIf OPX% = 0 Then GoTo 99
         End If

      ElseIf OPX% = 2 Then
         TIDOCUM$ = "Orden de Despacho"
         DOCUMENTO_ELEGIDO$ = TIDOCUM$
      ElseIf OPX% = 4 Then
        If InStr(1, EMPREAC$, "FERVI") < 1 Then
            Call COMUNI("Opción No Disponible")
            GoTo 10
        End If
        FILTX$ = TRIM$(NPX&)
        ARCSQR$ = "ODES-FAI"
        Call STDFORM(ARCSQR$, FILTX$)
        '
        GoTo 10
      ElseIf OPX% = 5 Then
         If DERACCE%("Refrescar Pedidos", "REFRPEDI") < 2 Then
           GoTo 10
         End If
         '
         'Call SELECHO("INDIPEDI/Indice General de Pedidos (Pendientes-Cumplidos-Anulados)")
         NUPEDIDO1& = NPX&
         If NUPEDIDO1& < 1 Then GoTo 10
         REFRESCARPED% = 1
         Call OPPED07.CARGA_PEDIDO(NUPEDIDO1&)
         '
         MODPED07.Show 1
         REFRESCARPED% = 0
         '
         Call CIERRARCH("*.*")
         GoTo 10
      End If
      '
      While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
      NUDOCU$ = NUDOCU$ + " "
      Screen.MousePointer = 11
      '
      
      DOCUNU& = DOCID&(TIDOCUM$, Str$(NPX&), NUDOCU$)
      If DOCUNU& < 1 Then
           Screen.MousePointer = 1
           Call COMUNI("No se Encontró el Documento Impreso\Correspondiente a la " & DOCUMENTO_ELEGIDO$ & " .\  \Si se Trata de un Pedido Pendiente de Entrega\Puede Refrescarlo con la Opción\'Modificación de Pedido en Curso'.")
           On Error Resume Next
           OPPED07.Command46.SetFocus
           OPPED07.Show 1
           If Err Then
            On Error GoTo 0
           End If
           GoTo 99
      End If
      Call MUESTRADOC(DOCUNU&, "")
   
        '
'''      GoTo 99
'''
'''      For UI& = ULTREG&("PDOCLST") To 1 Step -1
'''        XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
'''        If InStr(XX$, TIDOCUM$) > 4 Then
'''          PPXX% = InStr(XX$, "Nro.")
'''          If PPXX% > 0 Then
'''            XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
'''          End If
'''          If InStr(XX$, NUDOCU$) > 4 Then
'''            DOCUNU& = CVS(Left$(XX$, 4))
'''            GoTo 5
'''          End If
'''        End If
'''      Next UI&
'''      '
'''      Screen.MousePointer = 1
'''      Call COMUNI("No se Encontró el Documento Impreso\Correspondiente a este Pedido.\  \Si se Trata de un Pedido Pendiente de Entrega\Puede Refrescarlo con la Opción\'Modificación de Pedido en Curso'.")
'''      On Error Resume Next
'''      OPPED07.Command46.SetFocus
'''      OPPED07.Show 1
'''      If Err Then
'''      On Error GoTo 0
'''      End If
'''      GoTo 99
'''      '
'''5     Call MUESTRADOC(DOCUNU&) 'PERMITE VER LOS DOCUMENTOS ANTIGUOS POR PANTALLA
'''    End If
    '
       GoTo 10
99     Screen.MousePointer = 1
       Command6.Enabled = True
    '
End Sub

Private Sub Command7_Click()
   Command7.Enabled = False
   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If InStr(EPAC$, "DOSIVA") > 0 Then
       '
       PREPAP07.Show 1
       Call Label7_Change
       '
   Else
       '
       Call ABREPEDCLI ' SELECCION DE PEDIDOS PENDIENTES
       ' ** PONE EN BLANCO EL CAMPO NUMEROSERIE DE LOS PEDIDOS PREEXISTENTES
       Screen.MousePointer = 11
'       SQLY$ = "UPDATE PEDDETA SET NUMEROSERIE = '' WHERE NUMEROSERIE IS NULL"
'       FLEXCONN.Execute (FILTSQL$(SQLY$))
       ' **
       SQLX$ = " SELECT DISTINCT PEDENCA.NroPed, PEDENCA.NroClie, PEDENCA.FechEmis, PEDENCA.NroOcom, PEDENCA.Referencia "
       SQLX$ = SQLX$ + " FROM PEDENCA "
       SQLX$ = SQLX$ + " INNER JOIN  PEDDETA WITH(NOLOCK) "
       SQLX$ = SQLX$ + " ON (PEDENCA.NroPed = PEDDETA.NroPed) "
       SQLX$ = SQLX$ + " WHERE PEDDETA.Status < 8  "   ' para que no este anulada ni cumplida
       SQLX$ = SQLX$ + " AND PEDDETA.CantEnt < (PEDDETA.CanPed - 0.005) "
       SQLX$ = SQLX$ + " AND PEDENCA.Status < 9 "
       SQLX$ = SQLX$ + " AND PEDENCA.CODIEMPR = " & CodiEmp% & " "
       SQLX$ = SQLX$ + " AND PEDDETA.CODIEMPR = " & CodiEmp% & " "
       SQLX$ = SQLX$ + " ORDER BY PEDENCA.NroPed DESC "
       'jandy sql
       Dim PediModifTemp As ADODB.Recordset
       Set PediModifTemp = FLEXCONN.Execute(FILTSQL$(SQLX$))       'New ADODB.Recordset
       '
       Call COPYSTRU("INDIPEP", "INDIANSE")
       '
       JJ& = 0
       With PediModifTemp
          Do While Not .EOF
             NRODES& = !NROPED
             TTXX$ = MKS$(NRODES&)
             TTXX$ = TTXX$ + SAFETEXT$(!REFERENCIA)
             JJ& = JJ& + 1
             Call GRAREG("INDIANSE", Left$(TTXX$, 52) + MKS$(REGENCA&), JJ&)
          .MoveNext
          Loop
       End With
       PediModifTemp.Close
       Set PediModifTemp = Nothing
       '
       Call SETULTREG("INDIANSE", JJ&)
       Call CIERRARCH("INDIANSE")
       Call FRESHECHO("INDIANSE")
       '
       Screen.MousePointer = 1
       ASINUMSE.Show 1
   End If
   
   Command7.Enabled = True
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   '
   Screen.MousePointer = 11
   FIN& = ULTREG("PEDDETA")
   NPEDA& = 0
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     PDC$ = REGLEIDO$("PEDDETA", U&)
     FEMIPED% = CVI(Mid$(PDC$, 5, 2))
     If FEMIPED% >= FECHANUM("01/11/05") Then
       If Asc(Mid$(PDC$, 27, 1)) <> 9 Then
         If Mid$(PDC$, 51, 1) <> "R" Then
           NPED& = CVS(PDC$)
           CAN = CVS(Mid$(PDC$, 79, 4))
           ABUS1$ = Left$(PDC$, 26)
           ABUS2$ = Mid$(PDC$, 28, 43)
           '
           RFF$ = RECFILT$("DESPAPED", 1, MKS$(NPED&))
           For KU& = 1 To Len(RFF$) Step 4
             REDE& = CVS(Mid$(RFF$, KU&))
             XX$ = REGLEIDO$("DESPAPED", REDE&)
             If Left$(XX$, 26) = ABUS1$ Then
               If Mid$(XX$, 28, 43) = ABUS2$ Then
                 If CVS(Mid$(XX$, 75, 4)) = CAN Then
                   GoTo 90
                 End If
               End If
             End If
           Next KU&
           If NPED& <> NPEDA& Then
             Call MENSERR(24, "Pedido " & NPED& & " Inconsistente en DESPAPED")
             NPEDA& = NPED&
           End If
         End If
       End If
     End If
90 Next U&
   '
   Screen.MousePointer = 1
   Call COMUNI("Fin del Control")
   '
   Command8.Enabled = True
End Sub

Private Sub Command9_Click()
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Unload Me
   'Hide
End Sub

Private Sub Command43_Click()
  Command43.Enabled = False
  Call CIERRARCH("INDIPEDI") ' SE AGREGO ESTO POR QUE SI NO GENERABA UN ERROR 55
  EPAC$ = TRIM$(UCase$(EMPREAC$))
  If InStr(EPAC$, "TANIT") < 1 Then
    Call CARGALISEL("INDIPEDI", "PEDENCA", "NROPED/F6.0 ; REFERENCIA/A60", "NROPED > 0 ORDER BY NROPED DESC")
  Else
    Call CARGALISEL("INDIPEDI", "PEDENCA", "NROPED/F6.0 ; REFERENCIA/A60", "NROPED > 199000 ORDER BY NROPED DESC")
  End If
  Call FRESHECHO("INDIPEDI")
  Call CIERRARCH("INDIPEDI")
  Command43.Enabled = True
End Sub

Private Sub Form_Load()
    Call CENTRAFORM(Me)
    '
    DoEvents
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
       EMPRE_SIN_USUARIO$ = REPLACAR$(EPAC$, USERNAME$, "")
       Caption = Caption + "  -  " + EMPRE_SIN_USUARIO$
    End If    '
    '
    If InStr(EPAC$, "AHP") > 0 Then
       Command10.Visible = True
       Command26.Visible = True
    End If
    
    Screen.MousePointer = 1
    HOII% = HOY(HOS$)
    Label9 = HOS$
    Label12 = Left$(USERNAME, 17)
    '
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me, Picture3)
    '
    Unload FRMZELECHO

End Sub

Private Sub Label18_DblClick()
   NPX& = Val(Label7)
   If NPX& > 0 Then
     If DERACCE%("RESITPED", "Restaurar Situación de Preparacion y Despacho") = 2 Then
       If UCase$(Label14) = "VENTA" Then
         If COMALTER%("Realmente Desea Restaurar este Pedido\para Preparación y Despacho,\Revirtiendo TODOS los Movimientos ?\\No, Conservar\Si, Restaurar") = 2 Then
           '
           Screen.MousePointer = 11
           RFF$ = RECFILT$("DESPAPED", 1, MKS$(NPX&))
           If RFF$ = "" Then
             Call PREPAP07.RECUDESPA(NPX&)
             RFF$ = RECFILT$("DESPAPED", 1, MKS$(NPX&))
           End If
           '
           Screen.MousePointer = 11
           For U& = 1 To Len(RFF$) Step 4
             REPED& = CVS(Mid$(RFF$, U&, 4))
             XX$ = REGLEIDO$("DESPAPED", REPED&)
             Call REPLA(XX$, Chr$(1), 27, 1)
             Call GRAREG("DESPAPED", XX$, REPED&)
           Next U&
           Call CIERRARCH("DESPAPED")
           '
           REBLA$ = REGBLAN$("MOVISTO")
           ABUSBLA$ = "VP." + TRIM$(Str$(NPX&))
           ABUSBLB$ = "DM." + TRIM$(Str$(NPX&))
           FIN& = ULTREG&("MOVISTO")
           For KKU& = 1 To FIN
             Call TRACE(20, KKU&, FIN&)
             XX$ = REGLEIDO$("MOVISTO", KKU&)
             ABUSX$ = TRIM$(Mid$(XX$, 23, 10))
             If ABUSX$ = ABUSBLA$ Or ABUSX$ = ABUSBLB$ Then
               Call GRAREG("MOVISTO", REBLA$, KKU&)
               GoTo 19
             End If
             '
             ' ESTO ES PARA RESOLVER EL TEMA DEL PEDIDO 13260 DE DOSIVAC
             If ((Mid$(XX$, 21, 2) = "PP" Or Mid$(XX$, 21, 2) = "DM") And CVS(Mid$(XX$, 112, 4)) = NPX&) Then
               Call GRAREG("MOVISTO", REBLA$, KKU&)
               GoTo 19
             End If
19         Next KKU&
           Call CIERRARCH("MOVISTO")
           '
           Screen.MousePointer = 11
           TTXYZ$ = Space$(96)
           HOII% = HOY(HOS$)
           Call REPLA(TTXYZ$, HOS$, 1, 8)
           Call REPLA(TTXYZ$, "Restauracion de Preparados y Despacho ", 11, 56)
           Call REPLA(TTXYZ$, USERTER$, 69, 24)
 
           strSQL = "Select * from Pedenca where NroPed=" & NPX&
           Dim PedEnca As ADODB.Recordset
           Set PedEnca = New ADODB.Recordset
25         On Error Resume Next
           PedEnca.Open FILTSQL$(strSQL), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
           If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
               On Error GoTo 0
               Call CapturaErrorOpen
               GoTo 25
           End If
           On Error GoTo 0
           '
           ANOTANTE$ = ""
            '
           With PedEnca
              If Not (.EOF And .BOF) Then 'si no lo encuentra
                 ANOTA$ = TTXYZ$
                 ANOTA$ = ANOTA$ + Chr$(10) + Chr$(13)
                 ANOTA$ = ANOTA$ + String$(88, "=") + Chr$(10) + Chr$(13)
                 ANOTANTE$ = !AnotaPed
                 !AnotaPed = ANOTANTE$ + ANOTA$
                 .Update
              End If
           End With
           PedEnca.Close
           Set PedEnca = Nothing
           List1.Clear
           Label7 = "": Label7 = NPX&

           '\\\OT-08-0769-OD-FIN///
           Screen.MousePointer = 1
           Call COMUNI("Pedido " + TRIM$(Str$(NPX&)) + " Restaurado.\  \Recuerde que Debe Volver a Preparar\y Despachar los Materiales\Componentes de Este Pedido.")
         End If
       End If
     End If
   End If
End Sub

Private Sub Label7_Change()
   List1.Clear
   NPX& = XVALO(Label7)
   If NPX& < 1 Then Exit Sub
   '
20 Call APERBASDAT("PEDCLI")
   SQLX$ = "SELECT AnotaPed, RasoClie, NROOCOM, FechEmis, Status, DOMIENT, LOCAENT, TranspEnt, VenRep FROM PEDENCA WITH(NOLOCK) "
   SQLX$ = SQLX$ + " WHERE NroPed = " & NPX& & ""
'   Dim PEDDDENCA As New ADODB.Recordset
'   Set PEDDDENCA = New ADODB.Recordset
'25 On Error Resume Next
'   PEDDDENCA.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
'   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'      On Error GoTo 0
'      Call CapturaErrorOpen
'      GoTo 25
'   End If
'   On Error GoTo 0
   Set PEDDDENCA = READSET(SQLX$)
   '
   With PEDDDENCA
    On Error Resume Next
    .MoveFirst
    If Err < 1 Then
        On Error GoTo 0
        XX$ = SAFETEXT$(!AnotaPed): LONGI& = Len(XX$)
        Label11 = SAFETEXT$(!RaSoClie)
        Label16 = SAFETEXT$(!NROOCOM)
        FECEMI% = CVINT(!FECHEMIS)
        Label8 = FECHATEX$(FECEMI%)
        Label18 = SAFETEXT$(XVALO(!Status))
        LENTREGA(0) = SAFETEXT$(!DOMIENT)
        LENTREGA(1) = SAFETEXT$(!LOCAENT)
        LENTREGA(2) = SAFETEXT$(!TranspEnt)
        '
        Label14.ForeColor = RGB(0, 0, 255)
        Label14 = "Venta": REPARX% = 0
        If UCase$(Left$(SAFETEXT$(!VenRep), 1)) = "R" Then
            REPARX% = 1
            Label14.ForeColor = RGB(255, 0, 0)
            Label14 = "Reparación"
        End If
        '
        INICIO& = 1
        Do While INICIO& < LONGI&
            FIN& = InStr(INICIO&, XX$, Chr$(10))
            If FIN& > 0 Then
               TTXX$ = Mid$(XX$, INICIO&, FIN& - INICIO&)
               List1.AddItem TTXX$
               INICIO& = FIN& + 2
            Else
               TTXX$ = Mid$(XX$, INICIO&)
               List1.AddItem TTXX$
               Exit Do
            End If
        Loop
    End If
    On Error GoTo 0
   End With
   '
   ' ESTO ES PARA POSICIONARSE AL FINAL DE LA LISTA
   TPXI& = List1.ListCount - 15
   If TPXI& < 0 Then TPXI& = 0
   List1.TopIndex = TPXI&
   '
   List1.Refresh
   Openforms = DoEvents
   PEDDDENCA.Close
   Set PEDDDENCA = Nothing
   '
   ' 21/01/2006: al consultar un pedido que recupere
   '             la consistencia de preparacion y despacho
   If REPARX% <> 1 Then ' NO ES UNA REPARACIóN
     Screen.MousePointer = 11
     HUBOREC% = 0
     Call VERIDESPRE(NPX&, 0, HUBORECU%)
     If HUBORECU% > 0 Then HUBOREC% = 1
     Screen.MousePointer = 1
     If HUBOREC% > 0 Then
       Call COMUNI("Se Han Recuperado Acumulados\de Preparación y Despacho.")
     End If
   End If
   '
End Sub

Sub DESPAMAT()
   '
   Static DEPEX%
   '
   Screen.MousePointer = 11
   If DEPEX% < 1 Then
     DEPEX% = Int(Val(Control$("", "DEPOEXPE")))
     If DEPEX% < 1 Then
       Call MENSERR(24, "Proceso Imposible.\  \Falta Configurar Depósitos\de Expedición\y Productos Terminados.")
       GoTo 99
     End If
   End If
   '
   Screen.MousePointer = 11
   Call COPYSTRU("INDIPEDI", "PEPENDES")
   '
   Screen.MousePointer = 11
   SQLX$ = "SELECT DISTINCT PEDDETA.NroPed, PEDENCA.Referencia as REFEREN FROM PEDDETA WITH(NOLOCK)"
   SQLX$ = SQLX$ + " INNER JOIN PEDENCA"
   SQLX$ = SQLX$ + " ON PEDDETA.NroPed = PEDENCA.NroPed"
   SQLX$ = SQLX$ + " WHERE PEDDETA.NroPed > 0 "
   SQLX$ = SQLX$ + " AND PEDENCA.Status < 9 "                ' NO ANULADOS
   SQLX$ = SQLX$ + " AND PEDDETA.Status <> 8 "
   SQLX$ = SQLX$ + " AND PEDDETA.Status <> 9 "
   SQLX$ = SQLX$ + " AND PEDDETA.StaDespa < 2 "
   SQLX$ = SQLX$ + " AND PEDENCA.CODIEMPR = " & CodiEmp% & " "
   SQLX$ = SQLX$ + " AND PEDDETA.CODIEMPR = " & CodiEmp% & " "
   SQLX$ = SQLX$ + " AND AcuPrepa > AcuDespa"
   SQLX$ = SQLX$ + " AND PEDDETA.VenRep = 'V' "             ' Excluye Reparaciones
   SQLX$ = SQLX$ + " ORDER BY PEDDETA.NroPed ASC"
   Set IPEPE = New ADODB.Recordset
25 On Error Resume Next
   IPEPE.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
   '
   JJ& = 0
   With IPEPE
      On Error Resume Next
      .MoveLast
      If Err = 0 Then
         On Error GoTo 0
         FIN& = .RecordCount
         .MoveFirst
         Do While Not .EOF
           YY$ = MKS$(!NROPED) + !Referen
           JJ& = JJ& + 1
           Call TRACE(20, JJ&, FIN&)
           Call GRAREG("PEPENDES", YY$, JJ&)
         .MoveNext
         Loop
      End If
   End With
   '
   Call SETULTREG("PEPENDES", JJ&)
   Call FRESHECHO("PEPENDES")
   If JJ& < 1 Then
     Call COMUNI("No hay Pedidos Preparados\Pendientes de Despacho.")
     Screen.MousePointer = 1
     GoTo 99
   End If
   '
3  Screen.MousePointer = 1
   Call SELECHO("PEPENDES/Indice de Pedidos Pendientes de Despacho")
   NPX& = Val(VALACT1$("PEPENDES"))
   If NPX& < 1 Then GoTo 99
   '
5  ' Label7 = TRIM$(Str$(NPX&))
   '
   FORIPRE$ = TRIM$(Control$("", "CDESPACH"))
   If FORIPRE$ <> "" Then
     LU$ = LUDAT$
     If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
       RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
       Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Remito en Consignación:\  \'" + RMCC$ + "'.")
       Exit Sub
     End If
   End If
   '
   DEPOCON% = DEPOEXPE%
   DepOrig% = DEPOCON%
   DEPOSI$ = ECOARCH$("TADEPOSI", Chr$(DepOrig%))
   If DEPOCON% < 1 Or TRIM$(DEPOSI$) = "" Then
     Call MENSERR(24, "Falta Configurar Depósito de Expedición.")
     Exit Sub
   End If
   '
   Call CIERRARCH("*.*")
   Call BLANARCH("!BOLREDET")
   '
   Screen.MousePointer = 11
   NPED& = NPX&
   REBLA$ = REGBLAN$("DESPAPED")
   Call CIERRARCH("*.*")
   Call BLANARCH("!BOLREDET")
   '
   Screen.MousePointer = 11
   REBLA$ = REGBLAN$("DESPAPED")
   CONDI$ = " NroPed = " & NPX&
   CANTRE& = CantRegistros&("PEDDETA", CONDI$, "NOMESS")
   If CANTRE& < 1 Then
     Call MENSERR(24, "Imposible Despachar Pedido " + TRIM$(Str$(NPX&)))
     GoTo 3
   End If
   '
   SQLX$ = " SELECT CodiInt, Status, NroClie, CanPed, AcuPrepa, AcuDespa, FeSolEnt, NUORITEM  FROM PEDDETA WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE " & CONDI$  ' COINCIDE NÚMERO DE PEDIDO
   SQLX$ = SQLX$ + " order by NuOrItem "
   Set DESPAPE = READSET(SQLX$)
'   Dim DESPAPE As ADODB.Recordset
'   Set DESPAPE = New ADODB.Recordset
'26 On Error Resume Next
'   DESPAPE.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
'   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'      On Error GoTo 0
'      Call CapturaErrorOpen
'      GoTo 26
'   End If
'   On Error GoTo 0
   '
   With DESPAPE
'      On Error GoTo 0 'Resume Next
'      .MoveLast
'      If Err Then
'         On Error GoTo 0
'         Call MENSERR(24, "Imposible Despachar Pedido " + TRIM$(Str$(NPX&)))
'         GoTo 3
'      End If
'      On Error GoTo 0
'      .MoveFirst
      Do While Not .EOF
         CI% = XVALO(!CODIINT)
         STATU% = XVALO(!Status)
         NCLIE = XVALO(!NROCLIE)
         CAPENE# = XVALO(!CanPed)
         ACUPRE = XVALO(!AcuPrepa)
         ACUDES = XVALO(!AcuDespa)
         CAPENE# = ACUPRE - ACUDES
         FENTRESO% = CVINT(!FeSolEnt)
         NORITL& = XVALO(!NuOrItem)
         '
         If CAPENE# >= 0.05 Then
           Y$ = REGBLAN$("!BOLREDET")
           Call REPLA(Y$, MKI$(CI%), 83, 2)
           Call REPLA(Y$, Unimed$(CI%), 85, 4)
           Call REPLA(Y$, MKD$(CAPENE#), 89, 8)
           Call REPLA(Y$, MKS$(NPX&), 151, 4)
           Call REPLA(Y$, MKI$(FENTRESO%), 155, 2)
           Call REPLA(Y$, MKD$(CAPENE#), 163, 8)
           Call REPLA(Y$, MKS$(NORITL&), 189, 4)
           Call REGAPP("!BOLREDET", Y$)
         End If
      .MoveNext
      Loop
    End With
    DESPAPE.Close
    Set DESPAPE = Nothing
    '
12  VTB% = VENTABU%("DESPAPED/NC/TITUPAN=Despacho Materiales Vendidos - " + TRIM$(rasocli$(NCLIE)))
   If VTB% < 0 Or ULTREG&("!BOLREDET") < 1 Then
     Exit Sub
   End If
   '
   For YX& = 1 To ULTREG&("!BOLREDET")
     YY$ = REGLEIDO$("!BOLREDET", YX&)
     CI% = CVI(Mid$(YY$, 83, 2))
     CAN# = CVD(Mid$(YY$, 163, 8))
     If CI% > 0 Then
       If CI% <= NUMAS% Then
         If CAN# > 0.005 Then
           GoTo 13
         End If
       End If
     End If
   Next YX&
   Call MENSERR(24, "Despacho Cancelado !!!\Faltan Cantidades o Códigos.")
   Exit Sub
   '
13 QUEDAPEN% = 0
   For YX& = 1 To ULTREG&("!BOLREDET")
     YY$ = REGLEIDO$("!BOLREDET", YX&)
     CI% = CVI(Mid$(YY$, 83, 2))
     CAPREP# = CVD(Mid$(YY$, 163, 8))
     MAXX# = CVD(Mid$(YY$, 89, 8))
     If CAPREP# > MAXX# + 0.05 Then
       Call MENSERR(24, "Imposible Despachar Cantidad Mayor\a Preparada Pendiente de Despacho.")
       GoTo 12
     End If
     If CAPREP# < MAXX# - 0.05 Then QUEDAPEN% = 1
   Next YX&
   '
15 CONTA% = ULTREG&("!BOLREDET")
   If CONTA% > 0 Then
     Screen.MousePointer = 11  'Call BLIMESS("Registrando Recepcion ...")
     '
     TTXXX$ = "Materiales Vendidos"
     ANOTA$ = ""
16   TTXYZ$ = Space$(76)
     HOII% = HOY(HOS$)
     Call REPLA(TTXYZ$, HOS$, 1, 8)
     Call REPLA(TTXYZ$, "Despacho de " + TTXXX$ + ":", 11, 56)
     Call REPLA(TTXYZ$, USERTER$, 69, 24)
     ANOTA$ = ANOTA$ + TTXYZ$ + Chr$(10) + Chr$(13)
     ANOTA$ = ANOTA$ + Space$(88) + Chr$(10) + Chr$(13)
     '
     AINDI% = 0
     For YX& = 1 To ULTREG&("!BOLREDET")
       YY$ = REGLEIDO$("!BOLREDET", YX&)
       CI% = CVI(Mid$(YY$, 83, 2))
       CAN# = CVD(Mid$(YY$, 163, 8))
       LOTE$ = Mid$(YY$, 69, 12)
       If CI% > 0 Then
         If CI% <= NUMAS% Then
           If CAN# > 0.005 Then
             NUOC& = CVS(Mid$(YY$, 151, 4))
             NORII% = CVS(Mid$(YY$, 189, 4))
             CAVER$ = TRIM$(Str$(NPED&))
             While Len(CAVER$) > 6: CAVER$ = Mid$(CAVER$, 2): Wend
             While Len(CAVER$) < 6: CAVER$ = "0" + CAVER$: Wend
             IK$ = TRIM$(Str$(NORII%))
             CAVERI$ = CAVER$ + "-" + IK$
             DOSEC$ = CAVERI$
             CMO$ = "DM"
             REMI$ = "DM." + TRIM$(Str$(NPED&))
             FF% = HOY(HOS$)
             '
             CANTI = (-1) * CAN#
             Call MOVISTOK(FF%, CI%, LOTE$, CMO$, REMI$, REMI$, DOSEC$, DOSEC$, rasocli$(NCLIE), 0, VUNI#, 1, NCLIE, DepOrig%, CANTI, NPED&, NORII%)
             '
             TTXX$ = Space$(76)
             CIXI% = CI%
             CAENT = Abs(CANTI)
             Call REPLA(TTXX$, "- " + CODEXT$(CIXI%), 11, 18)
             Call REPLA(TTXX$, DESCRIT0$(CIXI%), 29, 38)
             Call REPLA(TTXX$, CSTRING$(MKS$(CAENT), 3, 10, 1, 2), 57, 10)
             ANOTA$ = ANOTA$ + Left$(TTXX$, 96) + Chr$(10) + Chr$(13)
             '
           End If
         End If
       End If
     Next YX&
     Call CIERRARCH("MOVISTO")
     '
     ANOTA$ = ANOTA$ + Left$(Space$(10) + String$(56, "-"), 96) + Chr$(10) + Chr$(13)
     ANOTA$ = ANOTA$ + String$(88, "=") + Chr$(10) + Chr$(13)
     '
     strSQL = "Select * from Pedenca where NroPed=" & NPX&
     Dim PedEnca As ADODB.Recordset
     Set PedEnca = New ADODB.Recordset
27   On Error Resume Next
     PedEnca.Open FILTSQL$(strSQL), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
     If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 27
     End If
     On Error GoTo 0
     ANOTANTE$ = ""
     '
     With PedEnca
       If Not (.EOF And .BOF) Then 'si no lo encuentra
          ANOTANTE$ = !AnotaPed
          !AnotaPed = ANOTANTE$ + ANOTA$
          .Update
       End If
     End With
     PedEnca.Close
     Set PedEnca = Nothing
     '
     Label7 = ""
     Label7 = TRIM$(Str$(NPED&))            ' PARA REFRESCAR
     '
     Screen.MousePointer = 1
     '
     If TRIM$(FORIPRE$) <> "" Then
       '
       FEX = FF%
       FECHDOC$ = FECHATEX$(FEX)
       NUMEDOC$ = TRIM$(REMI$)
       TIPODOC$ = "Despacho de Materiales"
       DESTIDOCNUE = NCLIE
       '
       CODPARAM$(1) = "UBI-MPRI": DOCPARAM$(1) = TRIM$(Label14)
       CAPARDOC% = 1
       '
       CODTABU1$(1) = "COD-MAT"
       CODTABU1$(2) = "DES-MAT"
       CODTABU1$(3) = "CANTIDAD"
       CODTABU1$(4) = "CAKILOS"
       CODTABU1$(5) = "UNIMED"
       CACOLTAB1% = 5
       '
       CAITAB1% = 0
       '
       For YX& = 1 To ULTREG&("!BOLREDET")
         YY$ = REGLEIDO$("!BOLREDET", YX&)
         CI% = CVI(Mid$(YY$, 83, 2))
         CAN# = CVD(Mid$(YY$, 89, 8))
         If CI% > 0 Then
           If CI% <= NUMAS% Then
             If CAN# > 0.005 Then
               UNID$ = TRIM$(Mid$(YY$, 85, 4))
               If UNID$ = "" Then UNID$ = Unimed$(CI%)
               CANS$ = CSTRING$(MKD$(CAN#), 3, 9, 1, 2)
               KILO$ = CSTRING$(MKD$(CAN# * PESUNI(CI%)), 3, 12, 3, 2)
               NULOTE$ = TRIM$(Mid$(YY$, 69, 12))
               NUOCO$ = CSTRING$(Mid$(YY$, 151, 4), 3, 7, 0, 2)
               '
               CAITAB1% = CAITAB1% + 1
               TETABU1$(1, CAITAB1%) = CODEXT$(CI%)
               TETABU1$(2, CAITAB1%) = DESCRIT$(CI%)
               TETABU1$(3, CAITAB1%) = CANS$
               TETABU1$(4, CAITAB1%) = KILO$
               TETABU1$(5, CAITAB1%) = UNID$
               '
             End If
           End If
         End If
       Next YX&
       '
       CANCELPRINT% = 0
       If Control$("CDESPACH", "SUPRINT") = "SI" Then CANCELPRINT% = 1
       Call PRINTDOC("CDESPACH")   ' Comprobante de Despacho
       '
     End If
   End If
   '
   If QUEDAPEN% < 1 Then
     Call FILTERDELETE("PEPENDES", 1, MKS$(NPED&))
     Call FRESHECHO("PEPENDES")
   End If
   GoTo 3
   '
99 Call CIERRARCH("*.*")
    
   '
End Sub


