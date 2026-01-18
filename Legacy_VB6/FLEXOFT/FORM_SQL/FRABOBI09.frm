VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FRABOBI09 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fracionamiento de Bobinas"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7275
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   2040
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   54
      Top             =   960
      Visible         =   0   'False
      Width           =   1170
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
      Left            =   6060
      TabIndex        =   37
      Top             =   315
      Width           =   175
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H00CDDADA&
      Caption         =   "FRACCIONAMIENTO"
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
      Left            =   240
      TabIndex        =   30
      Top             =   4920
      Width           =   6885
      Begin VB.CommandButton Command2 
         Caption         =   "O.K."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1180
         Left            =   5640
         Picture         =   "FRABOBI09.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Graba - Abrueba la Transacción"
         Top             =   360
         Width           =   945
      End
      Begin VB.Frame Frame18 
         Caption         =   "Frame3"
         Height          =   1260
         Left            =   3150
         TabIndex        =   46
         Top             =   315
         Width           =   15
      End
      Begin VB.Frame Frame17 
         Caption         =   "Frame14"
         Height          =   1260
         Left            =   1680
         TabIndex        =   45
         Top             =   315
         Width           =   15
      End
      Begin VB.Frame Frame7 
         Caption         =   "Frame3"
         Height          =   1260
         Left            =   2730
         TabIndex        =   44
         Top             =   315
         Width           =   15
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   930
         Left            =   210
         MousePointer    =   1  'Arrow
         TabIndex        =   42
         Top             =   600
         Width           =   5160
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Nro. Lote                Ancho(mm)  DP.  Depósito                                          Ubicación"
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
         Height          =   570
         Left            =   210
         TabIndex        =   43
         Top             =   345
         Width           =   5160
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00CDDADA&
      Caption         =   "MEDIDAS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   4320
      TabIndex        =   23
      Top             =   3045
      Width           =   1940
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   36
         Top             =   1935
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Kilos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   35
         Top             =   1000
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Ancho (mm)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   34
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label41 
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
         Height          =   285
         Left            =   220
         TabIndex        =   26
         Top             =   1260
         Width           =   1515
      End
      Begin VB.Label Label40 
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
         Height          =   285
         Left            =   220
         TabIndex        =   25
         Top             =   630
         Width           =   1515
      End
      Begin VB.Label Label39 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   600
         TabIndex        =   24
         Top             =   2250
         Visible         =   0   'False
         Width           =   1275
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Fechas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2430
      Left            =   7440
      TabIndex        =   19
      Top             =   720
      Visible         =   0   'False
      Width           =   2325
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Ingreso:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   105
         TabIndex        =   33
         Top             =   315
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Fabricación:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   105
         TabIndex        =   32
         Top             =   945
         Width           =   1335
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Vencimiento:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   105
         TabIndex        =   31
         Top             =   1575
         Width           =   1335
      End
      Begin VB.Label Label38 
         Alignment       =   2  'Center
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
         Height          =   285
         Left            =   630
         TabIndex        =   22
         Top             =   1890
         Width           =   1275
      End
      Begin VB.Label Label37 
         Alignment       =   2  'Center
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
         Height          =   285
         Left            =   630
         TabIndex        =   21
         Top             =   630
         Width           =   1275
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
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
         Height          =   285
         Left            =   630
         TabIndex        =   20
         Top             =   1260
         Width           =   1275
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00CDDADA&
      Caption         =   "IDENTIFICACIÓN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1830
      Left            =   240
      TabIndex        =   12
      Top             =   3045
      Width           =   3855
      Begin VB.Label Label1 
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
         Height          =   285
         Left            =   420
         TabIndex        =   52
         Top             =   795
         Width           =   480
      End
      Begin VB.Label Label20 
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
         Height          =   285
         Left            =   1140
         TabIndex        =   41
         Top             =   795
         Width           =   2400
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Depósito"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   240
         TabIndex        =   40
         Top             =   555
         Width           =   810
      End
      Begin VB.Label Label35 
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
         Height          =   285
         Left            =   420
         TabIndex        =   18
         Top             =   1395
         Width           =   3120
      End
      Begin VB.Label Label24 
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
         Height          =   285
         Left            =   420
         TabIndex        =   17
         Top             =   2205
         Visible         =   0   'False
         Width           =   3120
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Marca:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   105
         TabIndex        =   16
         Top             =   1965
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Lote de Fabricación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   30
         TabIndex        =   15
         Top             =   1140
         Width           =   1860
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "N.C.R."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   630
         TabIndex        =   14
         Top             =   390
         Width           =   810
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
         Height          =   285
         Left            =   1515
         TabIndex        =   13
         Top             =   315
         Width           =   2025
      End
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
      Height          =   285
      Left            =   3150
      TabIndex        =   11
      Top             =   315
      Width           =   175
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00CDDADA&
      Caption         =   "DETALLE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1710
      Left            =   240
      TabIndex        =   2
      Top             =   1260
      Width           =   6015
      Begin VB.Frame Frame4 
         Caption         =   "Frame4"
         Height          =   1215
         Left            =   4440
         TabIndex        =   50
         Top             =   360
         Width           =   15
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   1335
         Left            =   3360
         TabIndex        =   49
         Top             =   240
         Width           =   15
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   1335
         Left            =   2280
         TabIndex        =   48
         Top             =   240
         Width           =   15
      End
      Begin VB.Frame Frame13 
         Caption         =   "Frame3"
         Height          =   1170
         Left            =   0
         TabIndex        =   29
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame12 
         Caption         =   "Frame3"
         Height          =   1170
         Left            =   0
         TabIndex        =   28
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame11 
         Caption         =   "Frame3"
         Height          =   1170
         Left            =   0
         TabIndex        =   27
         Top             =   0
         Width           =   15
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1110
         ItemData        =   "FRABOBI09.frx":030A
         Left            =   210
         List            =   "FRABOBI09.frx":030C
         MousePointer    =   1  'Arrow
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   480
         Width           =   5565
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Kilos"
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
         Left            =   4485
         TabIndex        =   58
         Top             =   250
         Width           =   975
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ancho"
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
         Left            =   3405
         TabIndex        =   57
         Top             =   250
         Width           =   975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Ing."
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
         Left            =   2340
         TabIndex        =   56
         Top             =   250
         Width           =   1455
      End
      Begin VB.Label Label77 
         BackStyle       =   0  'Transparent
         Caption         =   "Lote"
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
         Left            =   260
         TabIndex        =   55
         Top             =   250
         Width           =   1455
      End
      Begin VB.Label Label25 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "      "
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   210
         TabIndex        =   5
         Top             =   240
         Width           =   5565
      End
   End
   Begin VB.PictureBox Picture14 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   4890
      Left            =   6360
      ScaleHeight     =   4860
      ScaleWidth      =   1350
      TabIndex        =   1
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
         TabIndex        =   59
         Top             =   4440
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   60
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command27 
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
         Height          =   650
         Left            =   120
         Picture         =   "FRABOBI09.frx":030E
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Otros Accesos Directos"
         Top             =   2520
         Width           =   650
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Kdx"
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
         Picture         =   "FRABOBI09.frx":0618
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Ficha Tipo Kardex de Movimientos Históricos"
         Top             =   3240
         Width           =   650
      End
      Begin VB.CommandButton Command3 
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
         Height          =   650
         Left            =   105
         Picture         =   "FRABOBI09.frx":0922
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   840
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
         Height          =   650
         Left            =   105
         Picture         =   "FRABOBI09.frx":0C2C
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Salir"
         Top             =   120
         Width           =   650
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   120
      OleObjectBlob   =   "FRABOBI09.frx":0D76
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   3000
      OleObjectBlob   =   "FRABOBI09.frx":0FAA
      TabIndex        =   53
      Top             =   1080
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label Label15 
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
      Height          =   300
      Left            =   3960
      TabIndex        =   39
      Top             =   400
      Width           =   945
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
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
      Height          =   285
      Left            =   5020
      TabIndex        =   38
      Top             =   315
      Width           =   1020
   End
   Begin VB.Label Label31 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   140
      TabIndex        =   9
      Top             =   810
      Width           =   1125
   End
   Begin VB.Label Label30 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Height          =   240
      Left            =   1335
      TabIndex        =   8
      Top             =   735
      Width           =   75
   End
   Begin VB.Label Label28 
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
      Height          =   285
      Left            =   1330
      TabIndex        =   7
      Top             =   315
      Width           =   1800
   End
   Begin VB.Label Label22 
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
      Height          =   300
      Left            =   -290
      TabIndex        =   6
      Top             =   400
      Width           =   1545
   End
End
Attribute VB_Name = "FRABOBI09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   '
   Call CIERRARCH("*.*")
   Call LIMPLAB
   Call LIMPTEXT
   Label28 = ""
   Label30 = ""
   List1.Clear
   Hide
   '
End Sub

Private Sub Command2_Click()

   Command2.Enabled = False
   '
   'VALIDACIONES
   FECHA% = FECHANUM(Label13)
   FEFA% = FECHANUM(Label36)
   FEVE% = FECHANUM(Label38)
   CIXI% = CODINT%(Label28)
   IDLOTX$ = TRIM$(Label2)
   DEPOO% = XVALO(Label1)
   '
   If DEPOO% < 1 Then
      Call MENSERR(24, "Falta Indicar Depósito de Origen")
      GoTo 99
   End If
   '
   If ECOARCH$("TADEPOSI", Chr$(DEPOO%)) = "" Then
      Call MENSERR(24, "Deposito Seleccionado no Valido")
      GoTo 99
   End If
   '
'   If XVALO(Label39) < 1 Then
'      Call MENSERR(24, "No Hay Cantidad")
'      GoTo 99
'   End If
   '
   If CIXI% < 1 Then
      Call MENSERR(24, "Codigo no Válido")
      GoTo 99
   End If
   '
   If IDLOTX$ = "" Then
      Call MENSERR(24, "Falta Seleccionar Lote a Fraccionar.")
      GoTo 99
   End If
   
   'Valida Total de Bobina con Suma de las Fracciones
   TOTFRA# = 0: ANCHO# = 0
   For i = 0 To List2.ListCount - 1
     ANCHO# = XVALO(Mid(List2.List(i), 17, 6))
     TOTFRA# = TOTFRA# + ANCHO#
   Next i
   '
   If Abs(TOTFRA# - XVALO(Label40)) >= 0.05 Then
      COMUNI ("Suma de Anchos Parciales\Difiere del Ancho Total")
      GoTo 99
   End If
   '
30  ' DETERMINA EL NÚMERO DE COMPROBANTE
   DOPRI$ = "CB-" + TRIM$(Str$(ProNroComprobante("CB")))
   DOSEC$ = DOPRI$
   REFEXI$ = ""
   For II& = 1 To List1.ListCount
      ANCHXX$ = " " + TRIM$(Mid$(List2.List(II& - 1), 17, 6))
      If InStr(REFEXI$, ANCHXX$) < 1 Then
         If REFEXI$ <> "" Then REFEXI$ = REFEXI$ + " /"
         REFEXI$ = REFEXI$ + ANCHXX$
      End If
   Next II&
   REFEXI$ = "Corte " + REFEXI$ + " mm"
   '
   VUNI# = 0
   MONEII% = 0
   NC% = 0
   'CANTI = (-1) * XVALO(Label39)
   CANTI = (-1) * XVALO(Label41)
   MARCAXX$ = Label24
   LOTORXX$ = Label35
   NCR$ = TRIM$(Label2)
   '
   ' declara movimiento de consumo lote original
   Call MOVISTOK(FECHA%, CIXI%, NCR$, "CB", DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFEXI$, 0, VUNI#, MONEII%, NC%, DEPOO%, CANTI)
   Call ACSALOTE(CIXI%, NCR$, "NOMESS")
   '
   For II& = 1 To List2.ListCount
      ' DECLARA MOVIMIENTO DE ENTRADA DEL PRIMER LOTE
      NCR1$ = Mid$(List2.List(II& - 1), 1, 15) 'NUM LOTE
      CANTLOTE$ = Mid$(List2.List(II& - 1), 17, 6)
      DEPO1$ = Mid$(List2.List(II& - 1), 25, 2)
      UBIFI1$ = TRIM$(Mid$(List2.List(II& - 1), 53, 24))
      REFE$ = "Fracc." + NCR$ + " (" + TRIM$(Label40) + " mm)"
     'CANTI = (XVALO(Label41) * XVALO(CANTLOTE$)) /1000
      CANTI = (XVALO(Label41) * XVALO(CANTLOTE$)) / XVALO(Label40)
      Call CREALOTIN(FECHA%, FEFA%, FEVE%, CIXI%, NCR1$, MARCAXX$, LOTORXX$, CDbl(CANTI), XVALO(Label41), XVALO(CANTLOTE$), XVALO(DEPO1$), UBIFI1$)
      Call MOVISTOK(FECHA%, CIXI%, NCR1$, "FB", DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, 0, VUNI#, MONEII%, NC%, XVALO(DEPO1$), CANTI)
      '
      CONDI$ = "IDENLOTE = '" & TRIM$(NCR$) & "'"    ' NCR$ ES LA BOBINA ORIGINAL SIN FRACCIONAR
      OPERAPRO$ = VALOBADA("LOTINGRE", "OperProd", CONDI$)
      TURNOPRO$ = VALOBADA("LOTINGRE", "TurnoProd", CONDI$)
      '
      CONDI$ = "IDENLOTE = '" & TRIM$(NCR1$) & "'"   ' NCR1$ ES LA BOBINA GENERADA POR EL FRACCIONAMIENTO
      Call ACTUREGISTRO("LOTINGRE", "OperProd", CONDI$, OPERAPRO$, REG&, "NOMESS")
      Call ACTUREGISTRO("LOTINGRE", "TurnoProd", CONDI$, TURNOPRO$, REG&, "NOMESS")
      '
   Next II&
   '
'   ' IMPRIME EN COMPROBANTE
'   For II& = 1 To List2.ListCount
'      NCR1$ = Mid$(List2.List(II& - 1), 1, 15) 'NUM LOTE
'      DEPO1$ = Mid$(List2.List(II& - 1), 25, 2)
'      Call PRINSTIKAPRO(CIXI%, NCR1$, XVALO(DEPO1$))
'   Next II&
'   '
   Call COMUNI("Transaccion Completada")
   Call Label28_Change
   Label39 = ""
   '
99 Screen.MousePointer = 1
   Command2.Enabled = True
   '
End Sub

Private Sub Command27_Click()
   '
'   OPX% = ALTERNA%("Control de Fraccionamiento\Otros Accesos Directos")
'   If OPX% = 1 Then
'     Call CONTROLFRACC
'   ElseIf OPX% = 2 Then
'     Call ACCDIR("CONSTOCK")
'   End If
    Call ACCDIR("CONSTOCK")

End Sub

Private Sub Command3_Click()
   '
   Command3.Enabled = False
   CUDI$ = UCase$(CurDir)
   PPXX% = InStr(CUDI$, "\WINDOWS")
   If PPXX% > 0 Then
      RUTAX$ = Left$(CUDI$, PPXX%) + "MANUALES\"
      If EXISTE%(RUTAX$ + "TRAZAB04.PDF") > 0 Then
         FORHLPDF.Pdf1.setShowToolbar (True)
         On Error Resume Next
         FORHLPDF.Pdf1.LOADFILE RUTAX$ + "TRAZAB04.pdf"
         If Err Then GoTo 20
         On Error GoTo 0
         FORHLPDF.Show 1
         GoTo 99
      End If
   End If
   '
20 Call MENSERR(24, "Imposible Abrir Archivo de Ayuda")
   '
99 Command3.Enabled = True
End Sub



Private Sub Command6_Click()
   '
   Call SELECHO("SELFECHA")
   VDEV$ = VALACT1$("SELFECHA")
   If VDEV$ <> "" Then
      Label13.Caption = VDEV$
   End If
   '
End Sub

Private Sub Command5_Click()
    '
    Call SELECHO("MASTER")
    If VALACT1$("MASTER") <> "" Then
       Label28 = VALACT1$("MASTER")
    End If
    '
End Sub

Private Sub Command8_Click()
    '
    Command8.Enabled = False
        FIKARDEX09.Show 1
    Command8.Enabled = True
    '
End Sub
'
Function CARPRODOF$(NROOF$)
  '
  CARPRODOF$ = ""
  Call ABREORFAB
  '
  SQLX$ = " SELECT COD_INTE FROM OPERFAB WITH(NOLOCK) "
  SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NROOF$ & "' "
'  Set OPFABTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
  Set OPFABTMP = READSET(SQLX$)
  '
  With OPFABTMP
    '
    On Error Resume Next
    .MoveFirst
    If Err Then
      On Error GoTo 0
      Exit Function
    End If
    On Error GoTo 0
    '
    CI% = XVALO(!cod_inte)
    CARPRODOF$ = CODEXT$(CI%)
  End With
  '
End Function



Private Sub Form_Load()
'    '
'    Call VERJOBID(OKEY%)
'    If OKEY% < 1 Then Call FINAL("")
'    '
'    EPAC$ = TRIM$(UCase$(EMPREAC$))
'    If EPAC$ <> "" Then
'       Caption = Caption + "  -  " + EPAC$
'    End If
'    Call GRATITFOR(Caption)
'    '
    Call APLICACIONSKINS(Me, Picture14)

    Call ABRECONEXION
    HOII% = HOY(HOS$)
    Label13 = HOS$
'    '
'    If Left$(APLINVO$, 7) = "BORECEP" Then ' Or INTERPRE% = 1 Then
'       Text1 = Mid$(APLINVO$, 8)
'       OPENFORMS = DoEvents
'    End If
'    '
End Sub

Sub VERIPERPRUE()
    '
    If Control$("TRAZABI", "FINICIO") = "" Then
       HOII% = HOY(HOS$)
       Call GRACONTROL("TRAZABI", "FINICIO", HOS$)
    End If
    '
    FINICIO = FECHANUM(Control$("TRAZABI", "FINICIO"))
    FECHACT = CVINT(FileDateTime(LUDAT$ + "MONIEXEC.DAT"))
    '
    Dim BB As String * 128
    NX1% = FILEOPEN(LUDAT$ + "MONIEXEC.DAT", 0, 128)
    FINI& = LOF(NX1%) / 128
    COMI& = FINI& - 128: If COMI& < 1 Then COMI& = 1
    '
    For U& = COMI& To FINI&
       Get #NX1%, U&, BB$
       FEJECU = FECHANUM(Mid$(BB$, 69, 8))
       If FEJECU > FECHACT Then
          FECHACT = FEJECU
       End If
    Next U&
    '
    If DIENTRE%(FINICIO, FECHACT) > 7 Then
       Call MENSERR(24, "Plazo de Verificación Vencido.")
       Call FINAL("")
    End If
    '
End Sub
'
Private Sub Form_Unload(Cancel As Integer)
    '
    Call CIERRARCH("*.*")
    Unload Me
    '
End Sub

Private Sub Frame15_DblClick()
    '
    If Label2 = "" Then Exit Sub
10  VTB% = VENTABU%("CARFRABO")
    If VTB% < 0 Then
      Call BLANARCH("!CARFRABO")
      List2.Clear: Exit Sub
    End If
    ANCHTOTA# = 0
    List2.Clear
    NCR$ = TRIM$(Label2)
    XXPO% = 65
    For i& = 1 To ULTREG&("!CARFRABO")
        ZZ$ = REGLEIDO$("!CARFRABO", i&)
        ANCHO# = CVS(Mid$(ZZ$, 3, 4))
        ANCHTOTA# = ANCHTOTA# + ANCHO#
        DEPOS% = CVI(Mid$(ZZ$, 7, 2))
        DEPOSITO$ = ECOARCH$("TADEPOSI", Chr$(DEPOS%))
        UBICACIO$ = TRIM$(Mid$(ZZ$, 9, 24))
        If ANCHO# < 0.05 Then
            Call MENSERR(24, "Falta Ingresar Ancho de Corte")
            GoTo 10
        End If
        If DEPOS% < 1 Then
            Call MENSERR(24, "Deposito No Valido")
            GoTo 10
        End If
'        If UBICACIO$ = "" Then
'            Call MENSERR(24, "Ubicacion Física no Válida")
'            GoTo 10
'        End If
        '
        Dim LOTINGTEMP As ADODB.Recordset
        PPXX% = InStr(NCR$, "/")
        If PPXX% > 0 Then NCR$ = Left$(NCR$, PPXX% - 1)
        For KKII% = XXPO% To 90
           NCRX$ = NCR$ + "/" + Chr$(KKII%)
           Set LOTINGTEMP = New ADODB.Recordset
           strSQL = "Select IdenLote FROM LOTINGRE WHERE IdenLote = '" & NCRX$ & "'"
           Set LOTINGTEMP = FLEXCONN.Execute(FILTSQL$(strSQL))
           If (LOTINGTEMP.EOF And LOTINGTEMP.BOF) Then
              NCR1$ = NCRX$
              NCR$ = NCR1$
              XXPO% = KKII% + 1
              GoTo 30
           End If
        Next KKII%
        
        Call MENSERR(24, "Imposible Declarar este Corte - Consulte")
        Exit Sub
        '
30      Z$ = Space$(128)
        Call REPLA(Z$, NCR1$, 1, 15)
        Call REPLA(Z$, FORMATNUM(ANCHO#, "F6.1"), 17, 6)
        Call REPLA(Z$, FORMATNUM(DEPOS%, "F2"), 25, 2)
        Call REPLA(Z$, DEPOSITO$, 29, 20)
        Call REPLA(Z$, UBICACIO$, 53, 24)
        List2.AddItem Z$
    Next i&
   
    ANCHOINI = XVALO(Label40.Caption)
    If ANCHOINI <> ROUND#(ANCHTOTA#, 1) Then
        Call MENSERR(24, "Ancho Total Fraccionado Difiere del Ancho Inicial")
        GoTo 10
    End If
    '
End Sub

Private Sub label1_Change()
   Label20 = ECOARCH$("TADEPOSI", Chr$(XVALO(Label1)))
End Sub


Private Sub Label24_Click()
    '
    Label17.Caption = Label24.Caption
    '
End Sub

Private Sub Label28_Change()
    '
    CIXI% = CODINT%(Label28)
    If CIXI% > 0 Then
        Label30 = DESCRIT0$(CIXI%)
        Call LIMPLAB
        Call LIMPTEXT
        Call ARMALILOTES(CIXI%)
    End If
    '
End Sub

Sub ARMALILOTES(CIXI%)
    '
     FLEXCONN.Execute (FILTSQL$("SET DATEFORMAT dmy"))
    '
    Call APERBASDAT("STOCKS")
    SQLX$ = " SELECT IdenLote, FECHALTA, AnchoMili, CanAlta, DepoUbic, UBICFIS FROM LOTINGRE WITH(NOLOCK)"
'    SQLX$ = SQLX$ + " INNER JOIN BOLRECEP"
'    SQLX$ = SQLX$ + " ON LOTINGRE.IDENLOTE = BOLRECEP.IDENLOTE"
    SQLX$ = SQLX$ + " WHERE LOTINGRE.Cod_Inte = " & CIXI%
'    SQLX$ = SQLX$ + " AND (CanAlta - CanConsu) > 0"    ' no trae los consumidos
    SQLX$ = SQLX$ + " AND cansaldo > 0"    ' no trae los consumidos
    SQLX$ = SQLX$ + " ORDER BY FechAlta ASC, LOTINGRE.IdenLote ASC "
    '
'    Dim LOTINGTEMP As ADODB.Recordset
'    Set LOTINGTEMP = New ADODB.Recordset
'    Set LOTINGTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    Set LOTINGTEMP = READSET(SQLX$)

    REBLA$ = Space$(256)
    JJ& = 0: List1.Clear
    With LOTINGTEMP
25     On Error Resume Next
       .MoveFirst
       If Err Then
         On Error GoTo 0
         Call MENSERR(24, "No Figuran Lotes Disponibles\para el Material Elegido.")
         GoTo 99
       End If
       '
       On Error GoTo 0
       Do While Not .EOF
         IDLOT$ = TRIM$(!idenlote)
         'ANCHOMM
         'Call ACSALOT(CIXI%, IDLOT$, CANTIN#, CACONRE#, CARESRE#, CSALDO#, NPROVX%)
         '
'         TTXX$ = REBLA$
'           FEX = CVINT(!FECHALTA)
'         Call REPLA(TTXX$, IDLOT$, 1, 16)
'         Call REPLA(TTXX$, FECHATEX$(FEX), 17, 8)
'           FEX = CVINT(!FechVenc)
'         Call REPLA(TTXX$, FECHATEX$(FEX), 27, 8)
'           ANCHX = !AnchoMili
'         Call REPLA(TTXX$, FORMATNUM$(ANCHX, "F10.1"), 35, 10)
'           CSALFI# = !CanSaldo: If CSALFI# < 0 Then CSALFI# = 0
'           METSAL = 0: If ANCHX > 0.1 Then METSAL = 1000 * CSALFI# / ANCHX
'         Call REPLA(TTXX$, FORMATNUM$(METSAL, "F10.0"), 45, 10)
'         Call REPLA(TTXX$, FORMATNUM$(CSALFI#, "F12,1"), 55, 12)
'           If !CanReser > 0.05 Then
'             Call REPLA(TTXX$, "SI", 69, 2)
'           End If
'           FEX = CVINT(!FechFab)
'         Call REPLA(TTXX$, FECHATEX$(FEX), 129, 8)
'         Call REPLA(TTXX$, MARCALOT$(CIXI%, IDLOT$), 137, 32)
'         Call REPLA(TTXX$, LOTORIG$(CIXI%, IDLOT$), 169, 24)
'         Call REPLA(TTXX$, FORMATNUM$(!DepoUbic, "F4.0"), 193, 4)
'         Call REPLA(TTXX$, SAFETEXT(!UBICFIS), 197, 32)
         '
         TTXX$ = REBLA$
         FEX = CVINT(!FechAlta)
         Call REPLA(TTXX$, IDLOT$, 1, 20)
         Call REPLA(TTXX$, FECHATEX$(FEX), 21, 8)
           ANCHX = !AnchoMili
         Call REPLA(TTXX$, FORMATNUM$(ANCHX, "F10.1"), 30, 10)
           KILOO = !CanAlta
         Call REPLA(TTXX$, FORMATNUM$(KILOO, "F10.1"), 40, 10)
         
         Call REPLA(TTXX$, MARCALOT$(CIXI%, IDLOT$), 137, 32)
         Call REPLA(TTXX$, LOTORIG$(CIXI%, IDLOT$), 169, 24)
         Call REPLA(TTXX$, FORMATNUM$(!DepoUbic, "F4.0"), 193, 4)
         Call REPLA(TTXX$, SAFETEXT(!UBICFIS), 197, 32)
         
         JJ& = JJ& + 1
         List1.AddItem TTXX$
       .MoveNext
       Loop
    End With
    '
99  LOTINGTEMP.Close
    Set LOTINGTEMP = Nothing
    '
End Sub

Private Sub Label28_DblClick()
    '
10  TTXX$ = InputBox$("Pinche sobre Barras", "Lectura Código de Barras")
    If Len(TTXX$) > 16 Then
       FRABOBI07.Label28 = TRIM$(Left$(TTXX$, 16))
       For KKI% = 1 To List1.ListCount
          LOTSE$ = TRIM$(Mid$(TTXX$, 17, 12))
          If TRIM$(Left$(List1.List(KKI% - 1), 12)) = LOTSE$ Then
             List1.Selected(KKI% - 1) = True
             Call List1_Click
             Exit Sub
          End If
       Next KKI%
       Call MENSERR(24, "Lectura no Valida")
       GoTo 10
    End If
    '
End Sub

Private Sub Label3_DblClick()
   Call Frame15_DblClick
End Sub


'
Private Sub List1_Click()
    
    Call LIMPTEXT
    '
    TTXX$ = List1.TEXT
    '
    NCR$ = TRIM$(Mid(TTXX$, 1, 12))
    COD$ = Label28
    CI% = CODINT%(COD$)
    Label2.Caption = NCR$
    Label24.Caption = TRIM$(Mid$(TTXX$, 137, 32)) ' MARCA
    Label35.Caption = TRIM$(Mid$(TTXX$, 169, 24)) ' LOTE DE FABRICACION
    Label37.Caption = TRIM$(Mid$(TTXX$, 17, 8))   ' INGRESO
    Label36.Caption = TRIM$(Mid$(TTXX$, 129, 8))  ' FABRICACION
    Label38.Caption = TRIM$(Mid$(TTXX$, 27, 8))   ' VENCIMIENTO
'    Label40.Caption = TRIM$(Mid$(TTXX$, 35, 10))  ' ANCHO
'    Label41.Caption = TRIM$(Mid$(TTXX$, 45, 10))  ' LARGO
'    Label39.Caption = TRIM$(Mid$(TTXX$, 55, 12))  ' CANTIDAD
    Label40.Caption = TRIM$(Mid$(TTXX$, 30, 10))  ' ANCHO
    Label41.Caption = TRIM$(Mid$(TTXX$, 40, 10))  ' kilos
    '
    Label16 = TRIM$(Mid$(TTXX$, 193, 4))
    Text5 = TRIM$(Mid$(TTXX$, 193, 4))
    Text4 = TRIM$(Mid$(TTXX$, 193, 4))
    Text2 = TRIM$(Mid$(TTXX$, 197, 32))
    Text3 = TRIM$(Mid$(TTXX$, 197, 32))
    Label1 = DEPOLOTE%(CI%, NCR$)
    'Call Text1_Change
    'Call MUESTRASEL
    '
End Sub

Sub LIMPLAB()
    '
    Label2.Caption = ""
    Label24.Caption = ""
    Label35.Caption = ""
    Label36.Caption = ""
    Label37.Caption = ""
    Label38.Caption = ""
    Label39.Caption = ""
    Label40.Caption = ""
    Label41.Caption = ""
    Label1.Caption = ""
    '
End Sub

Sub LIMPTEXT()
    Text4 = ""
    Text5 = ""
    Text8 = ""
    Text9 = ""
    Text14 = ""
    Text11 = ""
    Text16 = ""
    Text12 = ""
    Text17 = ""
    List2.Clear
    Call BLANARCH("!CARFRABO")
End Sub

Private Sub List1_DblClick()
    Call Frame15_DblClick
End Sub

Private Sub List2_DblClick()
   Call Frame15_DblClick
End Sub


Sub MUESTRASEL()
   '
   CIXI% = CODINT%(TRIM$(Label28))
   IDLOT$ = TRIM$(Label2)
   NUMDEP% = TRIM$(XVALO(Label1))
   Label39 = FORMATNUM$(SaldoLote(CIXI%, IDLOT$, NUMDEP%), "F10.1")
   ANCHX = XVALO(Label40.Caption) / 1000
   LARGX = 0: If ANCHX > 0 Then LARGX = XVALO(Label39) / ANCHX
   Label41.Caption = FORMATNUM$(LARGX, "F10.1")
   '
End Sub

Sub CONTROLFRACC()
   Screen.MousePointer = 11
   Call DESHASFECHA(Des%, Has%, PERIO$)
   If TRIM$(PERIO$) = "" Then GoTo 99
   FECDESDE$ = FETEXCOR1(Des%)
   FECHASTA$ = FETEXCOR1(Has%)
   
   INFORME$ = "Desde: " & FECDESDE$ & " --- HASTA: " & FECHASTA$ & vbCrLf & vbCrLf
   Mensaje$ = ""
   
   Call ABRECONEXION
'   Dim Movi As ADODB.Recordset
'   Set Movi = New ADODB.Recordset
   
   SQLX$ = "SELECT CANTSALID, DOPRILAR, COD_INTE FROM MOVISTO WITH(NOLOCK) WHERE CodiMovi = 'CB' AND FECHMOV >='" & FECDESDE$ & "' AND FECHMOV <='" & FECHASTA$ & "'"
'   Set Movi = FLEXCONN.Execute(FILTSQL$(SQLX$))
   Set Movi = READSET(SQLX$)

   With Movi
     .MoveFirst
     Do While Not .EOF
       SALIDA# = XVALO(!CANTSALID)
       DOCU$ = SAFETEXT(!DOPRILAR)
       Codigo$ = CODEXT(!cod_inte)
       
       CONDI$ = "DoPriLar ='" & DOCU$ & "' AND CODIMOVI ='FB' AND FECHMOV >='" & FECDESDE$ & "' AND FECHMOV <='" & FECHASTA$ & "'"
       TOTFB# = XVALO(VALOBADA("MOVISTO", "SUM(CantIngre)", CONDI$))
       If SALIDA# - TOTFB# > 0.05 Then
          'Call COMUNI("NO balancea fraccionamiento " & docu$ & " para codigo " & codigo$)
           'Diagnostico
           Mensaje$ = "NO balancea fraccionamiento " & DOCU$ & " para codigo " & Codigo$ & ",  Imp. Salida: " & SALIDA# & "  /  Tot. Frac: " & TOTFB#
           INFORME$ = INFORME$ + Mensaje$ + vbCrLf
       End If
       .MoveNext
     Loop
   End With
   Movi.Close
   Set Movi = Nothing
   Call SAVEFILE("C:\FLEXOFT\INFBOBIN.LOG", INFORME$)
   MsgBox ("Informe Guardado como C:\FLEXOFT\INFBOBIN.LOG")

   Call COMUNI("Proceso finalizado")
99 Screen.MousePointer = 1
End Sub
