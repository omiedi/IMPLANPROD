VERSION 5.00
Begin VB.Form CONUSE07 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Historial de Equipos"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   11745
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   9015
      Left            =   10800
      ScaleHeight     =   8985
      ScaleWidth      =   1170
      TabIndex        =   26
      Top             =   0
      Width           =   1200
      Begin VB.CommandButton Command7 
         Caption         =   "Reg. Armado"
         Default         =   -1  'True
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Registro de Armado"
         Top             =   5400
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   120
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   49
         Top             =   6840
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
      Begin VB.CommandButton Command22 
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
         Left            =   105
         Picture         =   "CONUSE07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Anotador de Pedidos"
         Top             =   3480
         Width           =   750
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
         Left            =   120
         Picture         =   "CONUSE07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Generar Orden de Fabricación"
         Top             =   4440
         Visible         =   0   'False
         Width           =   750
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
         Left            =   105
         Picture         =   "CONUSE07.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Sale de la Aplicación"
         Top             =   210
         Width           =   750
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
         Left            =   105
         Picture         =   "CONUSE07.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Configurar Comprobante de Despacho"
         Top             =   2160
         Visible         =   0   'False
         Width           =   750
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
         Left            =   105
         Picture         =   "CONUSE07.frx":0BA0
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Consultas y Re-Impresión"
         Top             =   2805
         Visible         =   0   'False
         Width           =   750
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
         Left            =   105
         Picture         =   "CONUSE07.frx":0EAA
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Graba las Modificaciones"
         Top             =   855
         Width           =   750
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
         Height          =   660
         Left            =   105
         Picture         =   "CONUSE07.frx":11B4
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   1500
         Width           =   750
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00CDDADA&
      Caption         =   "NUMERO DE SERIE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2265
      Left            =   120
      TabIndex        =   14
      Top             =   0
      Width           =   10590
      Begin VB.CommandButton Command8 
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
         Left            =   3600
         TabIndex        =   46
         Top             =   315
         Width           =   180
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
         Left            =   3600
         TabIndex        =   16
         Top             =   720
         Width           =   180
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Filtro por Cliente"
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
         Index           =   1
         Left            =   120
         TabIndex        =   48
         Top             =   405
         Width           =   1905
      End
      Begin VB.Label Label26 
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
         Left            =   2180
         TabIndex        =   47
         ToolTipText     =   "Doble Click para Limpiar Casillero de Filtro"
         Top             =   315
         Width           =   1410
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
         Height          =   255
         Left            =   5490
         TabIndex        =   44
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Tension"
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
         Left            =   4400
         TabIndex        =   43
         Top             =   1500
         Width           =   1275
      End
      Begin VB.Label Label23 
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
         Height          =   255
         Left            =   9330
         TabIndex        =   42
         Top             =   1440
         Width           =   975
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
         Height          =   255
         Left            =   6570
         TabIndex        =   41
         Top             =   1800
         Width           =   3735
      End
      Begin VB.Label Label21 
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
         Left            =   8000
         TabIndex        =   40
         Top             =   1500
         Width           =   1275
      End
      Begin VB.Label Label20 
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
         Height          =   255
         Left            =   1515
         TabIndex        =   39
         Top             =   1800
         Width           =   2295
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Factura"
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
         Left            =   -300
         TabIndex        =   38
         Top             =   1860
         Width           =   1695
      End
      Begin VB.Label Label18 
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
         Height          =   255
         Left            =   5490
         TabIndex        =   37
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "F.Entrega"
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
         Left            =   4120
         TabIndex        =   36
         Top             =   1860
         Width           =   1275
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
         Height          =   255
         Left            =   1515
         TabIndex        =   35
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "O/Fab.Nro"
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
         Left            =   -300
         TabIndex        =   34
         Top             =   1500
         Width           =   1695
      End
      Begin VB.Line Line8 
         X1              =   0
         X2              =   10560
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label8 
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
         Left            =   5400
         TabIndex        =   33
         Top             =   315
         Width           =   1065
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
         Left            =   4680
         TabIndex        =   21
         Top             =   720
         Width           =   5655
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
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
         Left            =   7035
         TabIndex        =   20
         Top             =   405
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
         Left            =   7875
         TabIndex        =   19
         Top             =   315
         Width           =   2460
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "F.Alta"
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
         Left            =   3960
         TabIndex        =   18
         Top             =   405
         Width           =   1275
      End
      Begin VB.Label Label7 
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
         Left            =   1515
         TabIndex        =   17
         Top             =   720
         Width           =   2070
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "N/Serie"
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
         Left            =   -480
         TabIndex        =   15
         Top             =   810
         Width           =   1905
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00CDDADA&
      Caption         =   "NOVEDADES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1950
      Left            =   120
      TabIndex        =   3
      Top             =   6200
      Width           =   10620
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
         Height          =   1175
         Left            =   1155
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
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
         TabIndex        =   10
         Top             =   1140
         Width           =   750
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   425
         Left            =   0
         ScaleHeight     =   390
         ScaleWidth      =   10545
         TabIndex        =   5
         Top             =   315
         Width           =   10580
         Begin VB.Line Line2 
            X1              =   0
            X2              =   1155
            Y1              =   420
            Y2              =   420
         End
         Begin VB.Label Label5 
            Caption         =   "Label5"
            Height          =   15
            Left            =   0
            TabIndex        =   7
            Top             =   420
            Width           =   1380
         End
         Begin VB.Label Label4 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha     Detalle                                                   Usuario"
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
            Left            =   105
            TabIndex        =   6
            Top             =   75
            Width           =   10200
         End
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
         Left            =   8400
         TabIndex        =   24
         Top             =   840
         Width           =   4000
      End
      Begin VB.Line Line6 
         X1              =   8295
         X2              =   10550
         Y1              =   1130
         Y2              =   1130
      End
      Begin VB.Label Label6 
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
         TabIndex        =   23
         Top             =   1310
         Width           =   960
      End
      Begin VB.Line Line5 
         X1              =   8295
         X2              =   8295
         Y1              =   735
         Y2              =   1900
      End
      Begin VB.Label Label9 
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
         Height          =   1165
         Left            =   0
         TabIndex        =   8
         Top             =   735
         Width           =   1170
      End
      Begin VB.Label Label10 
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
         Height          =   1170
         Left            =   8310
         TabIndex        =   9
         Top             =   720
         Width           =   2265
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CDDADA&
      Caption         =   "ANOTACIONES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3825
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   10620
      Begin VB.Frame Frame7 
         Caption         =   "Frame7"
         Height          =   3795
         Left            =   10290
         TabIndex        =   25
         Top             =   630
         Width           =   15
      End
      Begin VB.Frame Frame4 
         Caption         =   "Frame4"
         Height          =   4005
         Left            =   8085
         TabIndex        =   13
         Top             =   630
         Width           =   15
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   4110
         Left            =   1155
         TabIndex        =   12
         Top             =   630
         Width           =   15
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
         Height          =   2955
         Left            =   0
         TabIndex        =   11
         Top             =   735
         Width           =   10575
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   0
         ScaleHeight     =   405
         ScaleWidth      =   10545
         TabIndex        =   2
         Top             =   315
         Width           =   10580
         Begin VB.Label Label2 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha     Detalle                                                   Usuario"
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
            Left            =   120
            TabIndex        =   22
            Top             =   105
            Width           =   10200
         End
      End
      Begin VB.Line Line4 
         X1              =   0
         X2              =   10605
         Y1              =   4410
         Y2              =   4410
      End
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
      Width           =   1695
   End
End
Attribute VB_Name = "CONUSE07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TXTRAZA_CONUSE07$

Private Sub Command21_Click()

End Sub

Private Sub Command22_Click()
   '
   Screen.MousePointer = 11
   Command22.Enabled = False
   '
   NUSERI$ = TRIM$(Label7)
   If NUSERI$ = "" Then GoTo 99
   '
   CONDI$ = "NUMEROSERIE = '" & NUSERI$ & "'"
   NP& = XVALO(VALOBADA("PEDDETA", "NROPED", CONDI$, "NOMESS"))
   '
   If NP& < 1 Then  ' SI NO TIENE PEDIDO RELACIONADO
      Call MENSERR(24, "Número de Serie no Asignado a Pedido")
      GoTo 99
   End If
   '
   Screen.MousePointer = 1
   '
   On Error Resume Next
   CONPED07.Label7 = NP& ' PEGO EL NUMERO DE PEDIDO EN LABEL PRINCIPAL
   If Err Then
      On Error GoTo 0
      Call MENSERR(24, "Transacción no Válida Desde Esta Localización")
      GoTo 99
   End If
   On Error GoTo 0
   CONPED07.Command4.Enabled = False ' DESHABILITO BOTON DE SELECCION
   CONPED07.Show 1
   '
99 Screen.MousePointer = 1
   Command22.Enabled = True
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   NSERIE$ = TRIM$(Label7)
   If NSERIE$ <> "" Then
     Call ANOTANUSE(NSERIE$, Text1)
   End If
   Label7 = "": Label7 = NSERIE$
'     If TRIM$(Text1) <> "" Then
'       Call FRATEXTO(Text1, 56)
'
'       For KKI% = 1 To CARETEX%
'         TTX1$ = Space$(76)
'         Call REPLA(TTX1$, RETEX$(KKI%), 11, 56)
'         If KKI% = 1 Then
'           Call REPLA(TTX1$, HOS$, 1, 8)
'           Call REPLA(TTX1$, USERTER$, 69, 24)
'         End If
'         'Call REGAPP("ANOTAPED", Left$(MKS$(NPX&) + TTXX$, 96))
'         TTXX$ = TTXX$ + TTX1$ + Chr$(10) + Chr$(13)
'       Next KKI%
'       '
'       '\\\OT-08-0098-12/03/08-///
'       SQLX$ = "SELECT ANOTAPED FROM PEDENCA"
'       SQLX$ = SQLX$ + " WHERE NroPed = " & NPX& & ""
'       Dim PEDDDENCA As New ADODB.Recordset
'       Set PEDDDENCA = New ADODB.Recordset
'       PEDDDENCA.Open (SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
'       With PEDDDENCA
'          If Not (.BOF And .EOF) Then
'             ANOTA$ = !AnotaPed
'             ANOTA$ = ANOTA$ + TTXX$
'             ANOTA$ = ANOTA$ + String$(88, "=") + Chr$(10) + Chr$(13)
'             !AnotaPed = ANOTA$
'             .Update
'           End If
'       End With
'       PEDDDENCA.Close
'       Set PEDDDENCA = Nothing
'       '
'       Text1 = ""
'       '
'       List1.Clear
'       Label7 = "": Label7 = NPX&
'       '
'     End If
'   End If
   Command3.Enabled = True
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   If Label26 <> "" Then ' si tiene seleccionado algun cliente muestra el INDINUCL filtrado
   
   If ULTREG("INDINUCL") < 1 Then GoTo 99
    '
    Call SELECHO("INDINUCL")
    NSR$ = TRIM$(VALACT1$("INDINUCL"))
    If NSR$ = "" Then GoTo 99 ' si no selecciono ninguno sale
    '
    Label7 = NSR$
    GoTo 99
   End If
   '
   ' si no hay ningun cliente seleccionado como filtro muestra lista commpleta
   '
   Call SELECHO("INDINUSE")
   TTXX$ = TRIM$(VALACT1$("INDINUSE"))
   If TTXX$ <> "" Then Label7 = TRIM$(TTXX$)
   '
99 Command4.Enabled = True
End Sub

Private Sub Command5_Click()
'   Command5.Enabled = False
'   ACONE$ = "ORFABR07"
'   Call CONECRUN(ACONE$, "Ordenes de Fabricación")
'   Command5.Enabled = True
End Sub

Private Sub Command6_Click()
'   Command6.Enabled = False
'   NPX& = Val(Label7)
'   '
'   If NPX& > 0 Then
'      TIDOCUM$ = "Pedido de Cliente"
'      If DERACCE%("VERCOSTO/NOMESS", "Visualizar Costos y Precios") < 1 Then
'        Call MENSERR(24, "Usuario No Autorizado para Visualizar Pedidos.\La Consulta Cambia a Orden de Despacho")
'        TIDOCUM$ = "Orden de Despacho"
'      End If
'      '
'      NUDOCU$ = TRIM$(Str$(NPX&))
'      If TIDOCUM$ = "Pedido de Cliente" Then
'         If ALTERNA%("\Comprobante Impreso\Detalle de B.Datos") = 2 Then
'            ABUS$ = NUDOCU$
'            Call CONECRUN("GENREP07/DETAPEDNRO(" + ABUS$ + ")", "Consulta por Número de Remito")
'            GoTo 99
'         End If
'      End If
'
'      While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
'      NUDOCU$ = NUDOCU$ + " "
'      Screen.MousePointer = 11
'      '
'      For UI& = ULTREG&("PDOCLST") To 1 Step -1
'        XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
'        If InStr(XX$, TIDOCUM$) > 4 Then
'          PPXX% = InStr(XX$, "Nro.")
'          If PPXX% > 0 Then
'            XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
'          End If
'          If InStr(XX$, NUDOCU$) > 4 Then
'            DOCUNU& = CVS(Left$(XX$, 4))
'            GoTo 5
'          End If
'        End If
'      Next UI&
'      '
'      Screen.MousePointer = 1
'      Call COMUNI("No se Encontró el Documento Impreso\Correspondiente a este Pedido.\  \Si se Trata de un Pedido Pendiente de Entrega\Puede Refrescarlo con la Opción\'Modificación de Pedido en Curso'.")
'      On Error Resume Next
'      OPPED07.Command46.SetFocus
'      OPPED07.Show 1
'      If Err Then
'      On Error GoTo 0
'      End If
'      GoTo 99
'      '
'5     Call MUESTRADOC(DOCUNU&) 'PERMITE VER LOS DOCUMENTOS ANTIGUOS POR PANTALLA
'    End If
'    '
'99  Command6.Enabled = True
    '
End Sub

Sub BLANFORMU()
   
   List1.Clear
   '
   Label7 = "" ' NUMERO DE SERIE
   Label8 = "" ' FECHA
   Label11 = "" ' CODIGO EXTERNO
   Label14 = "" ' DESCRIPCION
   Label16 = "" 'ORDEN DE FABRICACION
   Label20 = "" ' FACTURA
   Label23 = ""
   Label22 = "" ' RAZON SOCIAL
   Label25 = "" ' TENSION
   Label18 = "" ' FECHENTREGA
   Text1 = ""   ' OBSERVACIONES
End Sub

Private Sub Command7_Click()
   NSERIE$ = TRIM$(Label7)
   If NSERIE$ = "" Then
     Call COMUNI("Debe Seleccionar Un N.Serie")
     Exit Sub
   End If
   FRMREGARMADO.Label13(2) = NSERIE$
   FRMREGARMADO.Show 1
End Sub

Private Sub Command8_Click()
   '
   Command8.Enabled = False
   '
   NCLIE = ListaSeleccionClientes("", True)
   If NCLIE < 1 Then GoTo 99
   '
   NC = NCLIE
   Label26 = TRIM$(Str$(NC))
   '
99 Screen.MousePointer = 1
   Command8.Enabled = True

End Sub

Private Sub Command9_Click()
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Hide
End Sub

Private Sub Form_Load()
    Call CENTRAFORM(Me)
    '
    DoEvents
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    EPAC$ = UCase$(EPAC$)
    If InStr(1, EPAC$, "EYM") > 0 Then
       Command7.Visible = True
    End If
    
    Screen.MousePointer = 1
    HOII% = HOY(HOS$)
    Label9 = " " + HOS$
    Label12 = Left$(USERNAME, 17)
    '
End Sub

Private Sub Label26_Change()
   '
   '
   NCLIE = XVALO(Label26)
   If NCLIE < 1 Then Exit Sub
   '
   Call BLANFORMU
TXTRAZA_CONUSE07$ = ""
   NC = NCLIE ' PRESENTA LOS DATOS DE NUMERO DE CLIENTE Y RAZON SOCIAL
   
   Label26 = TRIM$(Str$(NC))
   '
   Command4.Enabled = False ' BLOQUEO BOTON DE DESPLEGAR LISTA FILTRADA
   '                          '  HASTA QUE TERMINE EL PROCESO
TXTRAZA_CONUSE07$ = TXTRAZA_CONUSE07$ & "PUNTO 1 = " & Now & vbCrLf
   Call COPYSTRU("INDINUSE", "INDINUCL") '
   TTXX$ = REGBLAN$("INDINUCL")
   Call SETULTREG("!INDINUCL", 0)
TXTRAZA_CONUSE07$ = TXTRAZA_CONUSE07$ & "PUNTO 2 = " & Now & vbCrLf

   SQLX$ = "SELECT NUMEROSERIE, COD_INTE, Cod_Exte,DESCRIP  FROM TANUMSE WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE NUMECLI =" & NCLIE
   SQLX$ = SQLX$ + " ORDER BY Cod_Exte, NUMEROSERIE ASC "
'   Dim NUSERTMP As ADODB.Recordset
'   Set NUSERTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   Set NUSERTMP = READSET(SQLX$)
   '
TXTRAZA_CONUSE07$ = TXTRAZA_CONUSE07$ & "PUNTO 3 = " & Now & vbCrLf
   With NUSERTMP   ' CONSULTA PARA CARGAR EL LISTADO FILTRADO POR CLIENTE SELECCIONADO
     JJ& = 0
     Do While Not .EOF
      NUSERR$ = SAFETEXT$(!NUMEROSERIE)
      CI1% = XVALO(!COD_INTE)
      CODIX$ = TRIM$(SAFETEXT$(!COD_EXTE))
      DESXRIP$ = TRIM$(SAFETEXT$(!Descrip))
      Call REPLA(TTXX$, NUSERR$, 1, 16)
      Call REPLA(TTXX$, CODIX$ + "  " + DESXRIP$, 17, 47)
      JJ& = JJ& + 1
      Call GRAREG("!INDINUCL", TTXX$, JJ&) ' GRABA EL INDINUCL
      .MoveNext
     Loop
   End With
   NUSERTMP.Close
   Set NUSERTMP = Nothing
   '
   Call SETULTREG("!INDINUCL", JJ&)
   Call TRACENTRAL("INDINUCL", "")
   '
   If ULTREG&("INDINUCL") < 1 Then
     Call MENSERR(24, "No Hay Números de Serie\Para el Cliente Elegido")
     Command4.Enabled = True ' HABILITO BOTON DE SELECCION DE LSITADO FILTRADO
     Exit Sub
   End If
   '
   '
TXTRAZA_CONUSE07$ = TXTRAZA_CONUSE07$ & "PUNTO 4 = " & Now & vbCrLf
   Call FRESHECHO("INDINUCL")
   Command4.Enabled = True  'HABILITO BOTON DE SELECCION DE LSITADO FILTRADO

End Sub

Private Sub Label26_DblClick()
   Label26 = ""
End Sub

Private Sub Label7_Change()
   List1.Clear
   NSERIE$ = TRIM$(Label7)
   If NSERIE$ = "" Then Exit Sub
   '
TXTRAZA_CONUSE07$ = TXTRAZA_CONUSE07$ & "PUNTO 5 = " & Now & vbCrLf
   CONDI$ = "NUMEROSERIE = '" & NSERIE$ & " '"
   SQLX$ = "SELECT FECHALTA, COD_INTE, IdSubOr, NUMFACTURA, NUMECLI, TEN_ENSA, FechEntrega, HISTORIAL, Cod_Exte, Descrip "
   SQLX$ = SQLX$ + " FROM TANUMSE WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE " & CONDI$
   Set rs = READSET(SQLX$)
   With rs
        If Not .EOF Then
           FEC = FECHATEX(CVINT(rs!FECHALTA))  'FECHA DE ALTA
           Label8 = FEC
           CI% = XVALO(rs!COD_INTE) ' CODIGO
           Label11 = SAFETEXT$(rs!COD_EXTE)
           Label14 = SAFETEXT$(rs!Descrip) ' DESCRIPCION
           Label16 = SAFETEXT$(rs!IDSUBOR) 'ORDEN DE FABRICACION
           Label20 = SAFETEXT$(rs!NUMFACTURA) ' FACTURA
           NC = XVALO(rs!NUMECLI) ' CLIENTE
           Label23 = ""
           If NC > 0 Then
             Label23 = Str(NC)
           End If
           Label22 = rasocli$(NC)  ' RAZON SOCIAL
           Label25 = XVALO(rs!Ten_Ensa) ' TENSION
           FEC = FECHATEX(CVINT(rs!FechEntrega))  ' FECHENTR
           Label18 = FEC ' FECHENTREGA
           XX$ = SAFETEXT$(rs!Historial)
           rs.Close
         End If
   End With
   Set rs = Nothing
 TXTRAZA_CONUSE07$ = TXTRAZA_CONUSE07$ & "PUNTO 6 = " & Now & vbCrLf
'   FEC = FECHATEX(CVINT(VALOBADA("TANUMSE", "FECHALTA", CONDI$, "NOMESS")))  'FECHA DE ALTA
'   Label8 = FEC
'   CI% = XVALO(VALOBADA("TANUMSE", "COD_INTE", CONDI$, "NOMESS")) ' CODIGO
'   Label11 = CODEXT$(CI%)
'   Label14 = DESCRIT$(CI%) ' DESCRIPCION
'   Label16 = VALOBADA("TANUMSE", "IdSubOr", CONDI$, "NOMESS") 'ORDEN DE FABRICACION
'   Label20 = VALOBADA("TANUMSE", "NUMFACTURA", CONDI$, "NOMESS") ' FACTURA
'   NC = XVALO(VALOBADA("TANUMSE", "NUMECLI", CONDI$, "NOMESS")) ' CLIENTE
'   Label23 = ""
'   If NC > 0 Then
'     Label23 = Str(NC)
'   End If
'   Label22 = RASOCLI$(NC)  ' RAZON SOCIAL
'   Label25 = XVALO(VALOBADA("TANUMSE", "TEN_ENSA", CONDI$, "NOMESS")) ' TENSION
'   FEC = FECHATEX(CVINT(VALOBADA("TANUMSE", "FechEntrega", CONDI$, "NOMESS")))  ' FECHENTR
'   Label18 = FEC ' FECHENTREGA
'   XX$ = VALOBADA("TANUMSE", "HISTORIAL", CONDI$, "NOMESS")
   LONGI& = Len(XX$)
   
   INICIO& = 1
   Do While INICIO& < LONGI&
        FIN& = InStr(INICIO&, XX$, Chr$(13))
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
'GRABO EL ARCHIVO DE TRAZABILIDAD
For I1% = 1 To 32767
     ARCH$ = LUDAT$ & "TXTRAZA_CONUSE07-" & TRIM$(Str$(I1%)) & ".TXT"
     If EXISTE%(ARCH$) < 1 Then
       NARCHI2% = FILEOPEN%(ARCH$, 2, 128) ' ABRE EL ARCHIVO DE ESCRITURA
       Print #NARCHI2%, TXTRAZA_CONUSE07$ ' ESCRIBE EN EL ARCHIVO EL NUEVO TEXTO
       Close #NARCHI2%
       Exit For
     End If
Next I1%
   '
'20 Call APERBASDAT("PEDCLI")
'   SQLX$ = "SELECT * FROM PEDENCA"
'   SQLX$ = SQLX$ + " WHERE NroPed = " & NPX& & ""
'   Dim PEDDDENCA As New ADODB.Recordset
'   Set PEDDDENCA = New ADODB.Recordset
'   PEDDDENCA.Open (SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
'   '
'   With PEDDDENCA
'    On Error Resume Next
'    .MoveFirst
'    If Err < 1 Then
'        On Error GoTo 0
'        XX$ = !AnotaPed: LONGI& = Len(XX$)
'        Label11 = !RasoClie
'        Label16 = !NroOcom
'        FECEMI% = CVINT(!FechEmis)
'        Label8 = FECHATEX$(FECEMI%)
'        Label18 = !Status
'        LENTREGA(0) = !DOMIENT
'        LENTREGA(1) = !LOCAENT
'        LENTREGA(2) = !TranspEnt
'        '
'        Label14.ForeColor = RGB(0, 0, 255)
'        Label14 = "Venta": REPARX% = 0
'        If UCase$(Left$(SAFETEXT$(!VenRep), 1)) = "R" Then
'            REPARX% = 1
'            Label14.ForeColor = RGB(255, 0, 0)
'            Label14 = "Reparación"
'        End If
'        '
'        INICIO& = 1
'        Do While INICIO& < LONGI&
'            FIN& = InStr(INICIO&, XX$, Chr$(10))
'            If FIN& > 0 Then
'               TTXX$ = Mid$(XX$, INICIO&, FIN& - INICIO&)
'               List1.AddItem TTXX$
'               INICIO& = FIN& + 2
'            Else
'               TTXX$ = Mid$(XX$, INICIO&)
'               List1.AddItem TTXX$
'               Exit Do
'            End If
'        Loop
'    End If
'    On Error GoTo 0
'   End With
   '
   ' ESTO ES PARA POSICIONARSE AL FINAL DE LA LISTA
'   TPXI& = List1.ListCount - 15
'   If TPXI& < 0 Then TPXI& = 0
'   List1.TopIndex = TPXI&
'   '
'   List1.Refresh
'   OPENFORMS = DoEvents
'   PEDDDENCA.Close
'   Set PEDDDENCA = Nothing
'   '
'   ' 21/01/2006: al consultar un pedido que recupere
'   '             la consistencia de preparacion y despacho
'   If REPARX% <> 1 Then ' NO ES UNA REPARACIóN
'     Screen.MousePointer = 11
'     HUBOREC% = 0
'     Call VERIDESPRE(NPX&, 0, HUBORECU%)
'     If HUBORECU% > 0 Then HUBOREC% = 1
'     Screen.MousePointer = 1
'     If HUBOREC% > 0 Then
'       Call COMUNI("Se Han Recuperado Acumulados\de Preparación y Despacho.")
'     End If
'   End If
   '
End Sub

Sub DESPAMAT()
   '
   Static DEPEX%
   '
   Screen.MousePointer = 11
   If DEPEX% < 1 Then
     DEPEX% = Int(Val(CONTROL$("", "DEPOEXPE")))
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
   SQLX$ = SQLX$ + " AND AcuPrepa > AcuDespa"
   SQLX$ = SQLX$ + " AND PEDDETA.VenRep = 'V' "             ' Excluye Reparaciones
   SQLX$ = SQLX$ + " ORDER BY PEDDETA.NroPed ASC"
   Set IPEPE = New ADODB.Recordset
   IPEPE.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
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
   FORIPRE$ = TRIM$(CONTROL$("", "CDESPACH"))
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
   SQLX$ = " SELECT CodiInt, Status, NroClie, CanPed, AcuPrepa, AcuPrepa, AcuDespa, FeSolEnt, NUORITEM FROM PEDDETA "
   SQLX$ = SQLX$ + " WHERE NroPed = " & NPX& & " "   ' COINCIDE NÚMERO DE PEDIDO
   SQLX$ = SQLX$ + " order by NuOrItem "
   Dim DESPAPE As ADODB.Recordset
   Set DESPAPE = New ADODB.Recordset
   DESPAPE.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   '
   With DESPAPE
      On Error GoTo 0 'Resume Next
      .MoveLast
      If Err Then
         On Error GoTo 0
         Call MENSERR(24, "Imposible Despachar Pedido " + TRIM$(Str$(NPX&)))
         GoTo 3
      End If
      On Error GoTo 0
      .MoveFirst
      Do While Not .EOF
         CI% = !CODIINT
         STATU% = !Status
         NCLIE = XVALO(!NROCLIE)
         CAPENE# = XVALO(!CanPed)
         ACUPRE = XVALO(!AcuPrepa)
         ACUDES = XVALO(!AcuDespa)
         CAPENE# = ACUPRE - ACUDES
         FENTRESO% = CVINT(!FeSolEnt)
         NORITL& = !NuOrItem
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
     Anota$ = ""
16   TTXYZ$ = Space$(76)
     HOII% = HOY(HOS$)
     Call REPLA(TTXYZ$, HOS$, 1, 8)
     Call REPLA(TTXYZ$, "Despacho de " + TTXXX$ + ":", 11, 56)
     Call REPLA(TTXYZ$, USERTER$, 69, 24)
     Anota$ = Anota$ + TTXYZ$ + Chr$(10) + Chr$(13)
     Anota$ = Anota$ + Space$(88) + Chr$(10) + Chr$(13)
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
             Anota$ = Anota$ + Left$(TTXX$, 96) + Chr$(10) + Chr$(13)
             '
           End If
         End If
       End If
     Next YX&
     Call CIERRARCH("MOVISTO")
     '
     Anota$ = Anota$ + Left$(Space$(10) + String$(56, "-"), 96) + Chr$(10) + Chr$(13)
     Anota$ = Anota$ + String$(88, "=") + Chr$(10) + Chr$(13)
     '
     strsql = "Select AnotaPed Pedenca where NroPed=" & NPX&
     Dim PedEnca As ADODB.Recordset
     Set PedEnca = New ADODB.Recordset
     PedEnca.Open FILTSQL$(strsql), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
     ANOTANTE$ = ""
     '
     With PedEnca
       If Not (.EOF And .BOF) Then 'si no lo encuentra
          ANOTANTE$ = !AnotaPed
          !AnotaPed = ANOTANTE$ + Anota$
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
       If CONTROL$("CDESPACH", "SUPRINT") = "SI" Then CANCELPRINT% = 1
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



Private Sub Label7_DblClick()
   Call Command4_Click ' PRESENTA LISTA DE SELECCION DE TANUMSE
End Sub

