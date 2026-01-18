VERSION 5.00
Begin VB.Form FRABOBI07 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fracionamiento de Bobinas"
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   9345
   StartUpPosition =   2  'CenterScreen
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
      Left            =   7980
      TabIndex        =   46
      Top             =   315
      Width           =   175
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Fraccionamiento"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   105
      TabIndex        =   35
      Top             =   5565
      Width           =   8205
      Begin VB.Frame Frame19 
         Caption         =   "Frame14"
         Height          =   1260
         Left            =   1980
         TabIndex        =   61
         Top             =   315
         Width           =   15
      End
      Begin VB.Frame Frame18 
         Caption         =   "Frame3"
         Height          =   1260
         Left            =   3330
         TabIndex        =   58
         Top             =   315
         Width           =   15
      End
      Begin VB.Frame Frame17 
         Caption         =   "Frame14"
         Height          =   1260
         Left            =   1560
         TabIndex        =   57
         Top             =   315
         Width           =   15
      End
      Begin VB.Frame Frame16 
         Caption         =   "Frame3"
         Height          =   1260
         Left            =   5610
         TabIndex        =   56
         Top             =   315
         Width           =   15
      End
      Begin VB.Frame Frame7 
         Caption         =   "Frame3"
         Height          =   1260
         Left            =   3000
         TabIndex        =   55
         Top             =   315
         Width           =   15
      End
      Begin VB.ListBox List2 
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
         Height          =   960
         Left            =   90
         MousePointer    =   1  'Arrow
         TabIndex        =   53
         Top             =   630
         Width           =   7995
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Nro. Lote                  Cant   Ancho(mm)    DP.    Depósito                                    Ubicación"
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
         Left            =   90
         TabIndex        =   54
         Top             =   345
         Width           =   7995
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Medidas"
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
      Left            =   5985
      TabIndex        =   27
      Top             =   3045
      Width           =   2325
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad:"
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
         Left            =   0
         TabIndex        =   41
         Top             =   1575
         Width           =   915
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Largo (Mt):"
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
         Left            =   105
         TabIndex        =   40
         Top             =   945
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Ancho (mm):"
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
         Left            =   105
         TabIndex        =   39
         Top             =   315
         Width           =   1335
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
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
         Left            =   630
         TabIndex        =   30
         Top             =   1260
         Width           =   1275
      End
      Begin VB.Label Label40 
         Alignment       =   1  'Right Justify
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
         Left            =   630
         TabIndex        =   29
         Top             =   630
         Width           =   1275
      End
      Begin VB.Label Label39 
         Alignment       =   1  'Right Justify
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
         Left            =   630
         TabIndex        =   28
         Top             =   1890
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
      Left            =   3570
      TabIndex        =   23
      Top             =   3045
      Width           =   2325
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Ingreso:"
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
         Left            =   105
         TabIndex        =   38
         Top             =   315
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Fabricación:"
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
         Left            =   105
         TabIndex        =   37
         Top             =   945
         Width           =   1335
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Vencimiento:"
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
         Left            =   105
         TabIndex        =   36
         Top             =   1575
         Width           =   1335
      End
      Begin VB.Label Label38 
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
         Left            =   630
         TabIndex        =   26
         Top             =   1890
         Width           =   1275
      End
      Begin VB.Label Label37 
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
         Left            =   630
         TabIndex        =   25
         Top             =   630
         Width           =   1275
      End
      Begin VB.Label Label36 
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
         Left            =   630
         TabIndex        =   24
         Top             =   1260
         Width           =   1275
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Identificación"
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
      Left            =   105
      TabIndex        =   16
      Top             =   3045
      Width           =   3375
      Begin VB.CommandButton Command9 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   945
         TabIndex        =   51
         Top             =   800
         Width           =   175
      End
      Begin VB.TextBox Text6 
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
         IMEMode         =   3  'DISABLE
         Left            =   420
         TabIndex        =   50
         Top             =   800
         Width           =   540
      End
      Begin VB.Label Label20 
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
         TabIndex        =   52
         Top             =   800
         Width           =   1800
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dp0:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -200
         TabIndex        =   49
         Top             =   550
         Width           =   810
      End
      Begin VB.Label Label35 
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
         Left            =   420
         TabIndex        =   22
         Top             =   1995
         Width           =   2640
      End
      Begin VB.Label Label24 
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
         Left            =   420
         TabIndex        =   21
         Top             =   1365
         Width           =   2640
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Marca:"
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
         Left            =   105
         TabIndex        =   20
         Top             =   1120
         Width           =   705
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Lote de Fabricación:"
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
         Left            =   105
         TabIndex        =   19
         Top             =   1740
         Width           =   1860
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "N.C.R.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   630
         TabIndex        =   18
         Top             =   315
         Width           =   810
      End
      Begin VB.Label Label2 
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
         Left            =   1515
         TabIndex        =   17
         Top             =   315
         Width           =   1545
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
      TabIndex        =   15
      Top             =   315
      Width           =   175
   End
   Begin VB.CommandButton Command2 
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
      Height          =   1580
      Left            =   8400
      Picture         =   "FRABOBI07.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Graba - Abrueba la Transacción"
      Top             =   5670
      Width           =   830
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Detalle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1710
      Left            =   105
      TabIndex        =   4
      Top             =   1260
      Width           =   8205
      Begin VB.Frame Frame4 
         Caption         =   "Frame3"
         Height          =   1220
         Left            =   2940
         TabIndex        =   45
         Top             =   315
         Width           =   15
      End
      Begin VB.Frame Frame6 
         Caption         =   "Frame3"
         Height          =   1220
         Left            =   5985
         TabIndex        =   44
         Top             =   315
         Width           =   15
      End
      Begin VB.Frame Frame5 
         Caption         =   "Frame3"
         Height          =   1220
         Left            =   7245
         TabIndex        =   43
         Top             =   315
         Width           =   15
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame3"
         Height          =   1220
         Left            =   4935
         TabIndex        =   42
         Top             =   315
         Width           =   15
      End
      Begin VB.Frame Frame14 
         Caption         =   "Frame14"
         Height          =   1220
         Left            =   1890
         TabIndex        =   34
         Top             =   315
         Width           =   15
      End
      Begin VB.Frame Frame13 
         Caption         =   "Frame3"
         Height          =   1170
         Left            =   0
         TabIndex        =   33
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame12 
         Caption         =   "Frame3"
         Height          =   1170
         Left            =   0
         TabIndex        =   32
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame11 
         Caption         =   "Frame3"
         Height          =   1170
         Left            =   0
         TabIndex        =   31
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   1220
         Left            =   3885
         TabIndex        =   8
         Top             =   315
         Width           =   15
      End
      Begin VB.ListBox List1 
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
         Height          =   960
         Left            =   210
         MousePointer    =   1  'Arrow
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   630
         Width           =   7785
      End
      Begin VB.Label Label25 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "IdenLote         Fecha I.  Vencim.    Ancho     Largo         M2    Res"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   210
         TabIndex        =   7
         Top             =   315
         Width           =   7785
      End
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00400000&
      Height          =   5490
      Left            =   8400
      ScaleHeight     =   5430
      ScaleWidth      =   1320
      TabIndex        =   1
      Top             =   -20
      Width           =   1380
      Begin VB.CommandButton Command14 
         Caption         =   "Setup"
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
         Picture         =   "FRABOBI07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   60
         ToolTipText     =   "Configuracion de Funcionamiento"
         Top             =   3330
         Width           =   650
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
         Height          =   690
         Left            =   120
         Picture         =   "FRABOBI07.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Otros Accesos Directos"
         Top             =   2550
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
         Height          =   690
         Left            =   105
         Picture         =   "FRABOBI07.frx":0A56
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Ficha Tipo Kardex de Movimientos Históricos"
         Top             =   1785
         Width           =   690
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
         Height          =   690
         Left            =   105
         Picture         =   "FRABOBI07.frx":0D60
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1050
         Width           =   690
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   630
         TabIndex        =   2
         Top             =   4725
         Width           =   690
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command1 
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
         Picture         =   "FRABOBI07.frx":106A
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Salir"
         Top             =   300
         Width           =   690
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "FRABOBI07.frx":11B4
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   5040
         Width           =   1515
      End
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
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
      Left            =   5355
      TabIndex        =   48
      Top             =   375
      Width           =   1545
   End
   Begin VB.Label Label13 
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
      Left            =   7035
      TabIndex        =   47
      Top             =   315
      Width           =   1020
   End
   Begin VB.Label Label31 
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
      Height          =   300
      Left            =   210
      TabIndex        =   12
      Top             =   795
      Width           =   1125
   End
   Begin VB.Label Label30 
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
      Left            =   1470
      TabIndex        =   11
      Top             =   735
      Width           =   6690
   End
   Begin VB.Label Label28 
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
      Left            =   1470
      TabIndex        =   10
      Top             =   315
      Width           =   1860
   End
   Begin VB.Label Label22 
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
      Height          =   300
      Left            =   -210
      TabIndex        =   9
      Top             =   375
      Width           =   1545
   End
End
Attribute VB_Name = "FRABOBI07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub Imprimir_ParteCortePapel(NRO$)
    CODPARAM$(1) = "NUME-COM"
    CODPARAM$(2) = "FECH-EMI"
    CODPARAM$(3) = "FECH-LAN"
    CODPARAM$(4) = "REF-MAT1"
    CODPARAM$(5) = "REF-MAT2"
    CODPARAM$(6) = "MEDIDA-1"
    CODPARAM$(7) = "MEDIDA-2"
    CODPARAM$(8) = "MEDIDA-3"
    CODPARAM$(9) = "REF-MAT6"
    CODPARAM$(10) = "USUARIO"
    CAPARDOC% = 10
    '
    DOCPARAM$(1) = NRO$ ' SE TIENE QUE DETERMINAR
    DOCPARAM$(2) = Label13
    DOCPARAM$(4) = ""
    DOCPARAM$(5) = Label30
    DOCPARAM$(6) = Label2
    LARGO$ = FORMATNUM$(Label41, "I7")
    DOCPARAM$(7) = Label40 & " mm x " & TRIM$(LARGO$) & " Mt" ' medida mm x largo
    DOCPARAM$(8) = VALOBADA("LOTINGRE", "UBICFIS", "COD_INTE = " & CODINT%(Label28))
    DOCPARAM$(9) = Label20
    DOCPARAM$(10) = USERNAME$
    '
    CODTABU1$(1) = "REF-MAT3"
    CODTABU1$(2) = "REF-MAT4"
    CODTABU1$(3) = "REF-MAT5"
    CACOLTAB1% = 3
    '
    CAITAB1% = 0
    For i& = 1 To ULTREG&("!CARFRABO")
        CAITAB1% = CAITAB1% + 1
        zz$ = REGLEIDO$("!CARFRABO", i&)
        ancho# = CVS(Mid$(zz$, 3, 4))
        TETABU1$(1, CAITAB1%) = CVI(Mid$(zz$, 1, 2))
        TETABU1$(2, CAITAB1%) = ancho#
        TETABU1$(3, CAITAB1%) = Label41
    Next i&
    '
    FORMULARIO$ = CONTROL$("", "FORPACOR")
    Call PRINTDOC("@" + FORMULARIO$)   'PARTE DE CORTE DE PAPEL
    '
End Sub

Private Sub Command1_Click()
   '
   Call CIERRARCH("*.*")
   Call FINAL("")
   '
End Sub

Private Sub Command14_Click()
    OPTFRABOBI07.Show 1
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
   DEPO0% = XVALO(Text6)
   '
   If DEPO0% < 1 Then
      Call MENSERR(24, "Falta Indicar Depósito de Origen")
      GoTo 99
   End If
   '
   If XVALO(Label39) < 1 Then
      Call MENSERR(24, "No Hay Cantidad")
      GoTo 99
   End If
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
   TOTFRA# = 0: ancho# = 0
   For i = 0 To List2.ListCount - 1
     CANTBOB# = XVALO(Mid$(List2.List(i), 16, 3))
     ancho# = XVALO(Mid(List2.List(i), 19, 6))
     TOTFRA# = TOTFRA# + (ancho# * CANTBOB#)
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
   DEPORI% = XVALO(Text6)
   CANTI = (-1) * XVALO(Label39)
   MARCAXX$ = Label24
   LOTORXX$ = Label35
   NCR$ = TRIM$(Label2)
   '
   ' declara movimiento de consumo lote original
   Call MOVISTOK(FECHA%, CIXI%, NCR$, "CB", DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, 0, VUNI#, MONEII%, NC%, DEPORI%, CANTI)
   Call ACSALOTE(CIXI%, NCR$, "NOMESS")
   '
   For II& = 1 To List2.ListCount
    ' DECLARA MOVIMIENTO DE ENTRADA DEL PRIMER LOTE
    CANTBOB# = XVALO(Mid$(List2.List(II& - 1), 16, 3))
    For AX& = 1 To CANTBOB#
      NCR1$ = Mid$(List2.List(II& - 1), 1, 15) 'NUM LOTE
      CANTLOTE$ = Mid$(List2.List(II& - 1), 19, 6) 'CANTIDAD ANCHO
      DEPO1$ = Mid$(List2.List(II& - 1), 28, 2)
      UBIFI1$ = TRIM$(Mid$(List2.List(II& - 1), 54, 24))
      REFE$ = "Fracc." + NCR$ + " (" + TRIM$(Label40) + " mm)"
      CANTI = (XVALO(Label41) * XVALO(CANTLOTE$)) / 1000
      Call CREALOTIN(FECHA%, FEFA%, FEVE%, CIXI%, NCR1$, MARCAXX$, LOTORXX$, CDbl(CANTI), XVALO(Label41), XVALO(CANTLOTE$), XVALO(DEPO1$), UBIFI1$)
      Call MOVISTOK(FECHA%, CIXI%, NCR1$, "FB", DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, 0, VUNI#, MONEII%, NC%, XVALO(DEPO1$), CANTI)
    Next AX&
    '
   Next II&
   '
   ' IMPRIME EN COMPROBANTE
   For II& = 1 To List2.ListCount
      NCR1$ = Mid$(List2.List(II& - 1), 1, 15) 'NUM LOTE
      DEPO1$ = Mid$(List2.List(II& - 1), 25, 2)
      Call PRINSTIKAPRO(CIXI%, NCR1$, XVALO(DEPO1$))
   Next II&
   '
   Call Imprimir_ParteCortePapel(DOPRI$)
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
   OPX% = ALTERNA%("Control de Fraccionamiento\Otros Accesos Directos")
   If OPX% = 1 Then
     Call CONTROLFRACC
   ElseIf OPX% = 2 Then
     Call ACCDIR("CONSTOCK")
   End If


End Sub

Private Sub Command3_Click()
   '
   Command3.Enabled = False
   CUDI$ = UCase$(CurDir)
   PPXX% = InStr(CUDI$, "\WINDOWS")
   If PPXX% > 0 Then
      RUTAX$ = Left$(CUDI$, PPXX%) + "MANUALES\"
      If EXISTE%(RUTAX$ + "TRAZAB04.PDF") > 0 Then
'         FORHLPDF.Pdf1.setShowToolbar (True)
'         On Error Resume Next
'         FORHLPDF.Pdf1.LOADFILE rutax$ + "TRAZAB04.pdf"
'         If Err Then GoTo 20
'         On Error GoTo 0
'         FORHLPDF.Show 1
         GoTo 99
      End If
   End If
   '
20 Call MENSERR(24, "Imposible Abrir Archivo de Ayuda")
   '
99 Command3.Enabled = True
End Sub



Private Sub Command6_Click()

   Call SELECHO("SELFECHA")
   VDEV$ = VALACT1$("SELFECHA")
   If VDEV$ <> "" Then
      Text12.TEXT = VDEV$
   End If
   
End Sub

Private Sub Command5_Click()
    Call SELECHO("MASTER")
    If VALACT1$("MASTER") <> "" Then
       Label28 = VALACT1$("MASTER")
    End If
End Sub



Private Sub Command8_Click()
   Command8.Enabled = False
   Call TRAZORF07.Command10_Click
   Command8.Enabled = True
End Sub
'
Function CARPRODOF$(NROOF$)
  '
  CARPRODOF$ = ""
  Call ABREORFAB
  '
  SQLX$ = " SELECT * FROM OPERFAB "
  SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NROOF$ & "' "
  Set OPFABTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
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
    Ci% = !Cod_Inte
    CARPRODOF$ = CODEXT$(Ci%)
  End With
  '
End Function

Private Sub Command9_Click()
    Call CommSel("TADEPOSI", Command9, Text6)
End Sub

Private Sub Form_Load()
    '
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If EPAC$ <> "" Then
       Caption = Caption + "  -  " + EPAC$
    End If
    Call GRATITFOR(Caption)
    '
    Call APLICACIONSKINS(Me, Picture14)
    '
    Call ABRECONEXION
    HOII% = HOY(HOS$)
    Label13 = HOS$
    '
    If Left$(APLINVO$, 7) = "BORECEP" Then ' Or INTERPRE% = 1 Then
       Text1 = Mid$(APLINVO$, 8)
       Openforms = DoEvents
    End If
    '
End Sub

Sub VERIPERPRUE()
    '
    If CONTROL$("TRAZABI", "FINICIO") = "" Then
       HOII% = HOY(HOS$)
       Call GRACONTROL("TRAZABI", "FINICIO", HOS$)
    End If
    '
    FINICIO = FECHANUM(CONTROL$("TRAZABI", "FINICIO"))
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
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Frame15_DblClick()
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
        zz$ = REGLEIDO$("!CARFRABO", i&)
        ancho# = CVS(Mid$(zz$, 3, 4))
'        ANCHTOTA# = ANCHTOTA# + ANCHO#
        CANTBOB# = CVI(Mid$(zz$, 1, 2))
        ANCHTOTA# = ANCHTOTA# + (ancho# * CANTBOB#)
        DEPOS% = CVI(Mid$(zz$, 7, 2))
        Deposito$ = ECOARCH$("TADEPOSI", Chr$(DEPOS%))
        UBICACIO$ = TRIM$(Mid$(zz$, 9, 24))
        If ancho# < 0.05 Then
            Call MENSERR(24, "Falta Ingresar Ancho de Corte")
            GoTo 10
        End If
        If DEPOS% < 1 Then
            Call MENSERR(24, "Deposito No Valido")
            GoTo 10
        End If
        If UBICACIO$ = "" Then
            Call MENSERR(24, "Ubicacion Física no Válida")
            GoTo 10
        End If
        '
        For JJ = 1 To CANTBOB#
            Dim LOTINGTEMP As ADODB.Recordset
            PPXX% = InStr(NCR$, "/")
            If PPXX% > 0 Then NCR$ = Left$(NCR$, PPXX% - 1)
            For KKII% = XXPO% To 90
               NCRX$ = NCR$ + "/" + Chr$(KKII%)
               Set LOTINGTEMP = New ADODB.Recordset
               strsql = "Select IdenLote FROM LOTINGRE WHERE IdenLote = '" & NCRX$ & "'"
               Set LOTINGTEMP = FLEXCONN.Execute(FILTSQL(strsql))
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
30          Z$ = Space$(128)
    '        Call REPLA(Z$, NCR1$, 1, 15)
    '        Call REPLA(Z$, FORMATNUM(ANCHO#, "F6.1"), 17, 6)
    '        Call REPLA(Z$, FORMATNUM(DEPOS%, "F2"), 25, 2)
    '        Call REPLA(Z$, DEPOSITO$, 29, 20)
    '        Call REPLA(Z$, UBICACIO$, 53, 24)
            Call REPLA(Z$, NCR1$, 1, 15)
            Call REPLA(Z$, "1", 16, 3)
            Call REPLA(Z$, FORMATNUM(ancho#, "F6.1"), 19, 6)
            Call REPLA(Z$, FORMATNUM(DEPOS%, "F2"), 28, 2)
            Call REPLA(Z$, Deposito$, 32, 20)
            Call REPLA(Z$, UBICACIO$, 54, 24)
            List2.AddItem Z$
        Next JJ
    Next i&
    
    ANCHOINI = XVALO(Label40.Caption)
    If ANCHOINI <> ANCHTOTA# Then
        Call MENSERR(24, "Ancho Total Fraccionado Difiere del Ancho Inicial")
        GoTo 10
    End If
End Sub

Private Sub Label24_Click()

    Label17.Caption = Label24.Caption
    
End Sub

Private Sub Label28_Change()
    CIXI% = CODINT%(Label28)
    Label30 = DESCRIT0$(CIXI%)
    Call LIMPLAB
    Call LIMPTEXT
    Call ARMALILOTES(CIXI%)
End Sub

Sub ARMALILOTES(CIXI%)
    '
    FLEXCONN.Execute ("SET DATEFORMAT dmy")
    '
    Call APERBASDAT("STOCKS")
    SQLX$ = " SELECT * FROM LOTINGRE WITH(NOLOCK)"
'    SQLX$ = SQLX$ + " INNER JOIN BOLRECEP"
'    SQLX$ = SQLX$ + " ON LOTINGRE.IDENLOTE = BOLRECEP.IDENLOTE"
    SQLX$ = SQLX$ + " WHERE LOTINGRE.Cod_Inte = " & CIXI%
    SQLX$ = SQLX$ + " AND (CanAlta - CanConsu) > 0"    ' no trae los consumidos
    SQLX$ = SQLX$ + " ORDER BY FechAlta ASC, LOTINGRE.IdenLote ASC "
    '
    Dim LOTINGTEMP As ADODB.Recordset
    Set LOTINGTEMP = New ADODB.Recordset
    Set LOTINGTEMP = FLEXCONN.Execute(FILTSQL(SQLX$))
    '
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
         IDLOT$ = TRIM$(!IdenLote)
         'ANCHOMM
         'Call ACSALOT(CIXI%, IDLOT$, CANTIN#, CACONRE#, CARESRE#, CSALDO#, NPROVX%)
         '
         TTXX$ = REBLA$
           fex = CVINT(!FechAlta)
         Call REPLA(TTXX$, IDLOT$, 1, 16)
         Call REPLA(TTXX$, FECHATEX$(fex), 17, 8)
           fex = CVINT(!FechVenc)
         Call REPLA(TTXX$, FECHATEX$(fex), 27, 8)
           ANCHX = !AnchoMili
         Call REPLA(TTXX$, FORMATNUM$(ANCHX, "F10.1"), 35, 10)
           CSALFI# = !CanSaldo: If CSALFI# < 0 Then CSALFI# = 0
           METSAL = 0: If ANCHX > 0.1 Then METSAL = 1000 * CSALFI# / ANCHX
         Call REPLA(TTXX$, FORMATNUM$(METSAL, "F10.0"), 45, 10)
         Call REPLA(TTXX$, FORMATNUM$(CSALFI#, "F12,1"), 55, 12)
           If !CanReser > 0.05 Then
             Call REPLA(TTXX$, "SI", 69, 2)
           End If
           fex = CVINT(!FechFab)
         Call REPLA(TTXX$, FECHATEX$(fex), 129, 8)
         Call REPLA(TTXX$, MARCALOT$(CIXI%, IDLOT$), 137, 32)
         Call REPLA(TTXX$, LOTORIG$(CIXI%, IDLOT$), 169, 24)
         Call REPLA(TTXX$, FORMATNUM$(!DepoUbic, "F4.0"), 193, 4)
         Call REPLA(TTXX$, SAFETEXT(!UbicFis), 197, 32)
         '
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
    NCR$ = Mid(TTXX$, 1, 12)
    COD$ = Label28
    Ci% = CODINT%(COD$)
    Label2.Caption = NCR$
    Label24.Caption = TRIM$(Mid$(TTXX$, 137, 32)) ' MARCA
    Label35.Caption = TRIM$(Mid$(TTXX$, 169, 24)) ' LOTE DE FABRICACION
    Label37.Caption = TRIM$(Mid$(TTXX$, 17, 8))   ' INGRESO
    Label36.Caption = TRIM$(Mid$(TTXX$, 129, 8))  ' FABRICACION
    Label38.Caption = TRIM$(Mid$(TTXX$, 27, 8))   ' VENCIMIENTO
    Label40.Caption = TRIM$(Mid$(TTXX$, 35, 10))  ' ANCHO
    Label41.Caption = TRIM$(Mid$(TTXX$, 45, 10))  ' LARGO
    Label39.Caption = TRIM$(Mid$(TTXX$, 55, 12))  ' CANTIDAD
    '
    Label16 = TRIM$(Mid$(TTXX$, 193, 4))
    Text5 = TRIM$(Mid$(TTXX$, 193, 4))
    Text4 = TRIM$(Mid$(TTXX$, 193, 4))
    Text2 = TRIM$(Mid$(TTXX$, 197, 32))
    Text3 = TRIM$(Mid$(TTXX$, 197, 32))
    'Call Text1_Change
    Call MUESTRASEL
    '
End Sub

Sub LIMPLAB()
    Label2.Caption = ""
    Label24.Caption = ""
    Label35.Caption = ""
    Label36.Caption = ""
    Label37.Caption = ""
    Label38.Caption = ""
    Label39.Caption = ""
    Label40.Caption = ""
    Label41.Caption = ""
End Sub

Sub LIMPTEXT()
    Text4 = ""
    Text5 = ""
    Text6 = ""
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

Private Sub Text6_Change()
   Label20 = ECOARCH$("TADEPOSI", Chr$(XVALO(Text6)))
   Call MUESTRASEL
End Sub

Private Sub Text6_DblClick()
   Call Command9_Click
End Sub

Sub MUESTRASEL()
   '
   CIXI% = CODINT%(TRIM$(Label28))
   IDLOT$ = TRIM$(Label2)
   NUMDEP% = TRIM$(XVALO(Text6))
   Label39 = FORMATNUM$(SaldoLote(CIXI%, IDLOT$, NUMDEP%), "F10.1")
   ANCHX = XVALO(Label40.Caption) / 1000
   LARGX = 0: If ANCHX > 0 Then LARGX = XVALO(Label39) / ANCHX
   Label41.Caption = FORMATNUM$(LARGX, "F10.1")
   '
End Sub

Sub CONTROLFRACC()
   Screen.MousePointer = 11
   Call DESHASFECHA(DES%, HAS%, PERIO$)
   If TRIM$(PERIO$) = "" Then GoTo 99
   FECDESDE$ = FETEXCOR1(DES%)
   FECHASTA$ = FETEXCOR1(HAS%)
   
   INFORME$ = "Desde: " & FECDESDE$ & " --- HASTA: " & FECHASTA$ & vbCrLf & vbCrLf
   MENSAJE$ = ""
   
   Call ABRECONEXION
   Dim Movi As ADODB.Recordset
   Set Movi = New ADODB.Recordset
   
   SQLX$ = "SELECT * FROM MOVISTO WHERE CodiMovi = 'CB' AND FECHMOV >='" & FECDESDE$ & "' AND FECHMOV <='" & FECHASTA$ & "'"
   Set Movi = FLEXCONN.Execute(FILTSQL(SQLX$))
   With Movi
     .MoveFirst
     Do While Not .EOF
       SALIDA# = XVALO(!Cantsalid)
       docu$ = SAFETEXT(!DoPriLar)
       Codigo$ = CODEXT(!Cod_Inte)
       
       CONDI$ = "DoPriLar ='" & docu$ & "' AND CODIMOVI ='FB' AND FECHMOV >='" & FECDESDE$ & "' AND FECHMOV <='" & FECHASTA$ & "'"
       TOTFB# = XVALO(VALOBADA("MOVISTO", "SUM(CantIngre)", CONDI$))
       If SALIDA# - TOTFB# > 0.05 Then
          'Call COMUNI("NO balancea fraccionamiento " & docu$ & " para codigo " & codigo$)
           'Diagnostico
           MENSAJE$ = "NO balancea fraccionamiento " & docu$ & " para codigo " & Codigo$ & ",  Imp. Salida: " & SALIDA# & "  /  Tot. Frac: " & TOTFB#
           INFORME$ = INFORME$ + MENSAJE$ + vbCrLf
       End If
       .MoveNext
     Loop
   End With
   Call SaveFile("C:\FLEXOFT\INFBOBIN.LOG", INFORME$)
   MsgBox ("Informe Guardado como C:\FLEXOFT\INFBOBIN.LOG")

   Call COMUNI("Proceso finalizado")
99 Screen.MousePointer = 1
End Sub
