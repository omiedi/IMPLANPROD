VERSION 5.00
Begin VB.Form AJUSLO07 
   BackColor       =   &H00C0EEC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ajuste de Existencias por Lote"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   9345
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0EEC0&
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
      TabIndex        =   53
      Top             =   3045
      Width           =   3375
      Begin VB.Label Label2 
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
         Left            =   1995
         TabIndex        =   60
         Top             =   1365
         Width           =   1020
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Disponible:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   840
         TabIndex        =   59
         Top             =   1080
         Width           =   1545
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
         Left            =   735
         TabIndex        =   58
         Top             =   1890
         Width           =   2325
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Envase:"
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
         TabIndex        =   57
         Top             =   1575
         Width           =   810
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Lote / Partida:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   420
         TabIndex        =   56
         Top             =   315
         Width           =   1335
      End
      Begin VB.Label Label18 
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
         Left            =   1470
         TabIndex        =   55
         Top             =   630
         Width           =   1545
      End
      Begin VB.Label Label35 
         Height          =   330
         Left            =   210
         TabIndex        =   54
         Top             =   1050
         Visible         =   0   'False
         Width           =   855
      End
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
      Height          =   1250
      Left            =   8400
      Picture         =   "AJUSLO07.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   52
      ToolTipText     =   "Graba - Aprueba la Transacción"
      Top             =   5640
      Width           =   830
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
      Left            =   7980
      TabIndex        =   32
      Top             =   315
      Width           =   175
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H00C0EEC0&
      Caption         =   "Ajuste de Lote"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1365
      Left            =   120
      TabIndex        =   24
      Top             =   5520
      Width           =   8205
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
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
         Left            =   6720
         TabIndex        =   51
         Top             =   360
         Width           =   1215
      End
      Begin VB.Frame Frame6 
         Caption         =   "Frame5"
         Height          =   1250
         Left            =   5470
         TabIndex        =   40
         Top             =   150
         Width           =   15
      End
      Begin VB.Frame Frame5 
         Caption         =   "Frame5"
         Height          =   1250
         Left            =   2735
         TabIndex        =   39
         Top             =   150
         Width           =   15
      End
      Begin VB.Label Label16 
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
         Left            =   6720
         TabIndex        =   61
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label7 
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
         Left            =   5760
         TabIndex        =   50
         Top             =   880
         Width           =   855
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Kilos:"
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
         Left            =   6120
         TabIndex        =   49
         Top             =   400
         Width           =   855
      End
      Begin VB.Label Label27 
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
         Left            =   3960
         TabIndex        =   48
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label Label26 
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
         Left            =   3960
         TabIndex        =   47
         Top             =   840
         Width           =   1275
      End
      Begin VB.Label Label23 
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
         Left            =   3000
         TabIndex        =   46
         Top             =   880
         Width           =   855
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Kilos:"
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
         Left            =   3360
         TabIndex        =   45
         Top             =   400
         Width           =   855
      End
      Begin VB.Label Label5 
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
         Left            =   1200
         TabIndex        =   44
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label Label4 
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
         Left            =   1200
         TabIndex        =   43
         Top             =   840
         Width           =   1275
      End
      Begin VB.Label Label3 
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
         Left            =   240
         TabIndex        =   42
         Top             =   880
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Kilos:"
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
         Left            =   600
         TabIndex        =   41
         Top             =   400
         Width           =   855
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00C0EEC0&
      Caption         =   "Cantidad"
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
      TabIndex        =   16
      Top             =   3045
      Width           =   2330
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ubicacion:"
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
         Left            =   0
         TabIndex        =   30
         Top             =   1575
         Width           =   1035
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Kilos:"
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
         TabIndex        =   29
         Top             =   945
         Width           =   1335
      End
      Begin VB.Label Label10 
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
         Left            =   105
         TabIndex        =   28
         Top             =   315
         Width           =   1335
      End
      Begin VB.Label Label41 
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
         Left            =   630
         TabIndex        =   19
         Top             =   1260
         Width           =   1275
      End
      Begin VB.Label Label40 
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
         Left            =   630
         TabIndex        =   18
         Top             =   630
         Width           =   1275
      End
      Begin VB.Label Label39 
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
         TabIndex        =   17
         Top             =   1890
         Width           =   1275
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0EEC0&
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
      TabIndex        =   12
      Top             =   3045
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
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   25
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   1260
         Width           =   1275
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
      Left            =   3600
      TabIndex        =   11
      Top             =   315
      Width           =   175
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0EEC0&
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
      TabIndex        =   3
      Top             =   1260
      Width           =   8205
      Begin VB.Frame Frame1 
         Caption         =   "Frame3"
         Height          =   1220
         Left            =   0
         TabIndex        =   35
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame4 
         Caption         =   "Frame3"
         Height          =   1220
         Left            =   3930
         TabIndex        =   31
         Top             =   360
         Width           =   15
      End
      Begin VB.Frame Frame14 
         Caption         =   "Frame14"
         Height          =   1220
         Left            =   2880
         TabIndex        =   23
         Top             =   315
         Width           =   15
      End
      Begin VB.Frame Frame13 
         Caption         =   "Frame3"
         Height          =   1170
         Left            =   6900
         TabIndex        =   22
         Top             =   360
         Width           =   15
      End
      Begin VB.Frame Frame12 
         Caption         =   "Frame3"
         Height          =   1170
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame11 
         Caption         =   "Frame3"
         Height          =   1170
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   1220
         Left            =   5760
         TabIndex        =   6
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
         Left            =   600
         MousePointer    =   1  'Arrow
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   600
         Width           =   6945
      End
      Begin VB.Label Label25 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Lote                  F.Venc.   Ubicacion             kilos"
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
         Left            =   600
         TabIndex        =   5
         Top             =   315
         Width           =   6945
      End
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00400000&
      Height          =   5490
      Left            =   8400
      ScaleHeight     =   5430
      ScaleWidth      =   1320
      TabIndex        =   0
      Top             =   -20
      Width           =   1380
      Begin VB.CommandButton Command11 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "AJUSLO07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Salir"
         Top             =   120
         Width           =   690
      End
      Begin VB.CommandButton Command10 
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
         Height          =   690
         Left            =   120
         Picture         =   "AJUSLO07.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   870
         Width           =   690
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Kdx"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "AJUSLO07.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Ficha Tipo Kardex de Movimientos Históricos"
         Top             =   1605
         Width           =   690
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   630
         TabIndex        =   1
         Top             =   4725
         Width           =   690
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   2
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -285
         Picture         =   "AJUSLO07.frx":0A68
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
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5355
      TabIndex        =   34
      Top             =   375
      Width           =   1545
   End
   Begin VB.Label Label13 
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
      Left            =   7035
      TabIndex        =   33
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
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   210
      TabIndex        =   10
      Top             =   795
      Width           =   1125
   End
   Begin VB.Label Label30 
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
      Left            =   1470
      TabIndex        =   9
      Top             =   735
      Width           =   6690
   End
   Begin VB.Label Label28 
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
      Left            =   1470
      TabIndex        =   8
      Top             =   315
      Width           =   2220
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
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   -210
      TabIndex        =   7
      Top             =   375
      Width           =   1545
   End
End
Attribute VB_Name = "AJUSLO07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   '
   Call CIERRARCH("*.*")
   Hide
   '
End Sub

Private Sub Command10_Click()
 '
   Command10.Enabled = False
   CUDI$ = UCase$(CurDir)
   PPXX% = InStr(CUDI$, "\WINDOWS")
   If PPXX% > 0 Then
      rutax$ = Left$(CUDI$, PPXX%) + "MANUALES\"
      If EXISTE%(rutax$ + "TRAZAB04.PDF") > 0 Then
         FORHLPDF.Pdf1.setShowToolbar (True)
         On Error Resume Next
         FORHLPDF.Pdf1.LOADFILE rutax$ + "TRAZAB04.pdf"
         If Err Then GoTo 20
         On Error GoTo 0
         FORHLPDF.Show 1
         GoTo 99
      End If
   End If
   '
20 Call MENSERR(24, "Imposible Abrir Archivo de Ayuda")
   '
99 Command10.Enabled = True
End Sub

Private Sub Command11_Click()
   '
   Call CIERRARCH("*.*")
   Unload Me
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
   IDLOTX$ = TRIM$(Label18)
   DEPO0% = XVALO(Text6)
   '
   If CIXI% < 1 Then
      Call MENSERR(24, "Codigo no Válido.")
      GoTo 99
   End If
   '
   If XVALO(Text1.TEXT) < 0 Then
      Call MENSERR(24, "Kilos Ajustados no Válido.")
      GoTo 99
   End If
   '
   CMOV$ = TRIM$(CONTROL$("", "CODAJUST"))
   If CMOV$ = "" Then CMOV$ = "AJ"
   
   DPRI$ = TRIM$(Str$(1 + XVALO(Mid$(VALOBADA("MOVISTO", "MAX(DOPRILAR)", "CodiMovi = '" & CMOV$ & "'"), 4, 6))))
   While Len(DPRI$) < 6: DPRI$ = "0" + DPRI$: Wend
   DPRI$ = CMOV$ + "-" + DPRI$
   HOII% = HOY(HOS$)
   DSEC$ = CMOV$ + "-" + Mid$(HOS$, 7, 2) + Mid$(HOS$, 4, 2) + Mid$(HOS$, 1, 2)
   For KKI% = 0 To 99
      KKX$ = TRIM$(Str$(KKI%))
      While Len(KKX$) < 2: KKX$ = "0" + KKX$: Wend
      KKY$ = VALOBADA("MOVISTO", "DOSECLAR", "DOSECLAR = '" + DSEC$ + "-" + KKX$ + "'")
      If KKY$ = "" Then
         DSEC$ = DSEC$ + "-" + KKX$
         GoTo 25
      End If
   Next KKI%
   Call MENSERR(24, "Imposible Procesar en Fecha Elegida - Cambie Fecha Movimiento.")
   GoTo 99
   '
25 Call ACSALOTE(CIXI%, IDLOTX$, "NOMESS")
   CONDI$ = "Cod_Inte = " & CIXI% & " "
   CONDI$ = CONDI$ + " AND IdenLote = '" & IDLOTX$ & "' "
   KILDISLO = XVALO(VALOBADA("LOTINGRE", "CanSaldo", CONDI$))
   If KILDISLO + XVALO(Label27) < 0 Then
       Call MENSERR(24, "Ajuste Imposible.\Excede Saldo Disponible del Lote Elegido.")
       GoTo 99
   End If
   '
   Call MOVISTOK(FECHA%, CIXI%, TRIM$(Label18), CMOV$, DPRI$, DPRI$, DSEC$, DSEC$, "Ajuste Existencias ", 0, 0, 0, 0, DEPX%, XVALO(Label27))
   Call ACSALOTE(CIXI%, TRIM$(Label18), "NOMESS")
   '
   Call LIMPLAB
   Call LIMPTEXT
   Call Label28_Change
 
99 Screen.MousePointer = 1
   Command2.Enabled = True
   '
End Sub

Private Sub Command3_Click()
   '
   Command3.Enabled = False
   CUDI$ = UCase$(CurDir)
   PPXX% = InStr(CUDI$, "\WINDOWS")
   If PPXX% > 0 Then
      rutax$ = Left$(CUDI$, PPXX%) + "MANUALES\"
      If EXISTE%(rutax$ + "TRAZAB04.PDF") > 0 Then
         FORHLPDF.Pdf1.setShowToolbar (True)
         On Error Resume Next
         FORHLPDF.Pdf1.LOADFILE rutax$ + "TRAZAB04.pdf"
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



Private Sub Command4_Click()
    Call SELECHO("TAUBICA")
    UBISELE$ = VALACT1$("TAUBICA")
    If UBISELE$ <> "" Then
       Text3 = UBISELE$
    End If
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

Private Sub Command7_Click()
   Command7.Enabled = False
   FIKARDEX07.Show 1
   Command7.Enabled = True
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   FIKARDEX07.Show 1
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
    CI% = !Cod_Inte
    CARPRODOF$ = CODEXT$(CI%)
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
    Call ABRECONEXION
    HOII% = HOY(HOS$)
    Label13 = HOS$
    '
    If Left$(APLINVO$, 7) = "BORECEP" Then ' Or INTERPRE% = 1 Then
       Text1 = Mid$(APLINVO$, 8)
       OPENFORMS = DoEvents
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
     FLEXCONN.EXECUTE (FILTSQL$ ("SET DATEFORMAT dmy"))
    '
    Call APERBASDAT("STOCKS")
    SQLX$ = " SELECT * FROM LOTINGRE WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE LOTINGRE.Cod_Inte = " & CIXI%
    'sqlx$ = sqlx$ + " AND (CanAlta - CanConsu) > 0"    ' SUPRIMIDO PARA PODER AJUSTAR TAMBIEN LOS CONSUMIDOS
    SQLX$ = SQLX$ + " ORDER BY FechAlta ASC, LOTINGRE.IdenLote ASC "
    '
    Dim LOTINGTEMP As ADODB.Recordset
    Set LOTINGTEMP = New ADODB.Recordset
    Set LOTINGTEMP =  FLEXCONN.EXECUTE (FILTSQL$ (SQLX$))
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
         '
         TTXX$ = REBLA$
         Call REPLA(TTXX$, IDLOT$, 1, 20) 'lote
         FEX = CVINT(!FechVenc)
         Call REPLA(TTXX$, FECHATEX$(FEX), 23, 8)                           ' Fecha vencimiento
         Call REPLA(TTXX$, SAFETEXT(!UbicFis), 33, 16)                      ' Ubicacion
         Call REPLA(TTXX$, FORMATNUM$(XVALO(!CanAlta - !CanConsu), "F12.1"), 48, 12)  ' Kilos
         Call REPLA(TTXX$, XVALO(!DepoUbic), 70, 2)                         ' Deposito
         Call REPLA(TTXX$, SAFETEXT(!CodEnvase), 72, 25)                    ' Envase
         FEXALTA = CVINT(!FechAlta)
         Call REPLA(TTXX$, FECHATEX$(FEXALTA), 97, 8)                       ' fecha Ingreso
         FEXFAB = CVINT(!FechFab)
         Call REPLA(TTXX$, FECHATEX$(FEXFAB), 105, 8)                       ' Fecha fabricación
            COEFI = 1: If !CanAlta > 0 Then COEFI = !Cantidad / !CanAlta
         Call REPLA(TTXX$, FORMATNUM(XVALO(COEFI * (!CanAlta - !CanConsu)), "F12.1"), 113, 12) ' Cantidad
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

Private Sub List1_Click()
    
    Call LIMPLAB
    Call LIMPTEXT
    '
    TTXX$ = List1.TEXT
    '
    NCR0$ = Mid(TTXX$, 1, 20)
    COD$ = Label28
    CI% = CODINT%(COD$)
    Label18.Caption = NCR0$
    Label24.Caption = TRIM$(Mid$(TTXX$, 72, 25)) ' Envase
    Label37.Caption = TRIM$(Mid$(TTXX$, 97, 8))   ' F. Ingreso
    Label36.Caption = TRIM$(Mid$(TTXX$, 105, 8))  ' F. Fabricación
    Label38.Caption = TRIM$(Mid$(TTXX$, 23, 8))   ' F. Vencimiento
    Label40.Caption = TRIM$(Mid$(TTXX$, 113, 12))  ' Cantidad
    Label41.Caption = TRIM$(Mid$(TTXX$, 49, 12))  ' Kilos
    Label39.Caption = TRIM$(Mid$(TTXX$, 33, 16))  ' Ubicación
    Label34 = Label39.Caption
    Label33 = Label40
    Label5 = Label41.Caption
    Label4 = Label40.Caption
    Label2 = TRIM$(VALOBADA("LOTINGRE", "Libera", "IdenLote = '" & NCR0$ & "'"))
       COEFI = 1: If XVALO(Label41) > 0 Then COEFI = XVALO(Label40) / XVALO(Label41)
    Label35 = FORMATNUM$(COEFI, "F15.5")

End Sub

Sub LIMPLAB()
    Label2.Caption = ""
    Label24.Caption = ""
    Label36.Caption = ""
    Label37.Caption = ""
    Label38.Caption = ""
    Label39.Caption = ""
    Label40.Caption = ""
    Label41.Caption = ""
    Label5.Caption = ""
    Label4.Caption = ""
    Label26.Caption = ""
    Label27.Caption = ""
    Label18.Caption = ""
End Sub

Sub LIMPTEXT()
    Text1 = ""
    Text2 = ""
    Text3 = ""
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
    Call BLANARCH("!CARFRABO")
End Sub

Private Sub Text1_Change()
   '
   Label27.Caption = FORMATNUM$(XVALO(Text1.TEXT) - XVALO(Label5.Caption), "F12.1") 'RESTA AL LOTE ORIGINAL LOS KILOS AJUSTADOS
   COEFI = XVALO(Label35)
   Label16 = FORMATNUM$(XVALO(Text1) * COEFI, "F12.1")
   Label26 = FORMATNUM$(XVALO(Label27) * COEFI, "F12.1")
   '
End Sub

Private Sub Text1_LostFocus()
   '
   Text1.TEXT = TRIM$(FORMATNUM$(Text1.TEXT, "F12.1")) ' FORMATEA EL TEXT1 "KILOS"
   '
End Sub

Private Sub Text2_Change()
   '
   Label26.Caption = TRIM$(FORMATNUM$(XVALO(Label4.Caption) - XVALO(Text2.TEXT), "F12.1"))  'RESTA AL LOTE ORIGINAL LA CANTIDAD AJUSTADA
   '
End Sub

Private Sub Text2_LostFocus()
   '
   Text2.TEXT = FORMATNUM$(Text2.TEXT, "F12.1") ' FORMATEA EL TEXT2 "CANTIDAD"
   '
End Sub

Private Sub Text6_Change()
   Label20 = ECOARCH$("TADEPOSI", Chr$(XVALO(Text6)))
End Sub

Private Sub Text6_DblClick()
   Call Command9_Click
End Sub

