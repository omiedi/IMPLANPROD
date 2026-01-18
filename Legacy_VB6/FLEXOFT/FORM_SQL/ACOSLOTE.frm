VERSION 5.00
Begin VB.Form ACOSLOTE 
   BackColor       =   &H00C0E0F0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actualizacion de Costos por Lote"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   9345
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0E0F0&
      Caption         =   "Costo de Reposición"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   105
      TabIndex        =   46
      Top             =   5145
      Width           =   8205
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
         Height          =   285
         Left            =   7815
         TabIndex        =   55
         Top             =   315
         Width           =   165
      End
      Begin VB.TextBox Text25 
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
         Left            =   6885
         TabIndex        =   53
         Text            =   " "
         Top             =   315
         Width           =   945
      End
      Begin VB.CommandButton Command15 
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
         Left            =   3780
         TabIndex        =   50
         Top             =   315
         Width           =   165
      End
      Begin VB.TextBox Text13 
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
         Left            =   3465
         MaxLength       =   2
         TabIndex        =   49
         Top             =   315
         Width           =   360
      End
      Begin VB.TextBox Text1 
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
         Height          =   285
         Left            =   1050
         TabIndex        =   47
         Top             =   315
         Width           =   1410
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6195
         TabIndex        =   54
         Top             =   375
         Width           =   645
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Moneda:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   8
         Left            =   2625
         TabIndex        =   52
         Top             =   375
         Width           =   765
      End
      Begin VB.Label Label3 
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
         Left            =   3990
         TabIndex        =   51
         Top             =   315
         Width           =   1830
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Valor:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   420
         TabIndex        =   48
         Top             =   375
         Width           =   645
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00C0E0F0&
      Caption         =   "Medidas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2430
      Left            =   5985
      TabIndex        =   26
      Top             =   2625
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
         TabIndex        =   39
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
         TabIndex        =   38
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
         Index           =   0
         Left            =   105
         TabIndex        =   37
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
         TabIndex        =   29
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
         TabIndex        =   28
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
         TabIndex        =   27
         Top             =   1890
         Width           =   1275
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0E0F0&
      Caption         =   "Fechas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2430
      Left            =   3570
      TabIndex        =   22
      Top             =   2625
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
         TabIndex        =   36
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
         TabIndex        =   35
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
         TabIndex        =   34
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
         Top             =   1260
         Width           =   1275
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0E0F0&
      Caption         =   "Identificación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2430
      Left            =   105
      TabIndex        =   15
      Top             =   2625
      Width           =   3375
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dp0: "
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
         Left            =   1680
         TabIndex        =   45
         Top             =   315
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.Label Label16 
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
         Left            =   2520
         TabIndex        =   44
         Top             =   210
         Visible         =   0   'False
         Width           =   495
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
         TabIndex        =   21
         Top             =   1890
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
         TabIndex        =   20
         Top             =   1155
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
         TabIndex        =   19
         Top             =   840
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
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   105
         TabIndex        =   18
         Top             =   1575
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
         Left            =   60
         TabIndex        =   17
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
         Left            =   1300
         TabIndex        =   16
         Top             =   630
         Width           =   1755
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
      Left            =   2205
      TabIndex        =   14
      Top             =   420
      Width           =   175
   End
   Begin VB.CommandButton Command2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1350
      Left            =   8400
      Picture         =   "ACOSLOTE.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Graba - Abrueba la Transacción"
      Top             =   4620
      Width           =   825
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0F0&
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
      Height          =   1710
      Left            =   105
      TabIndex        =   4
      Top             =   840
      Width           =   8205
      Begin VB.Frame Frame4 
         Caption         =   "Frame3"
         Height          =   1220
         Left            =   2940
         TabIndex        =   43
         Top             =   315
         Width           =   15
      End
      Begin VB.Frame Frame6 
         Caption         =   "Frame3"
         Height          =   1220
         Left            =   5985
         TabIndex        =   42
         Top             =   315
         Width           =   15
      End
      Begin VB.Frame Frame5 
         Caption         =   "Frame3"
         Height          =   1220
         Left            =   7245
         TabIndex        =   41
         Top             =   315
         Width           =   15
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame3"
         Height          =   1220
         Left            =   4935
         TabIndex        =   40
         Top             =   315
         Width           =   15
      End
      Begin VB.Frame Frame14 
         Caption         =   "Frame14"
         Height          =   1220
         Left            =   1890
         TabIndex        =   33
         Top             =   315
         Width           =   15
      End
      Begin VB.Frame Frame13 
         Caption         =   "Frame3"
         Height          =   1170
         Left            =   0
         TabIndex        =   32
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame12 
         Caption         =   "Frame3"
         Height          =   1170
         Left            =   0
         TabIndex        =   31
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame11 
         Caption         =   "Frame3"
         Height          =   1170
         Left            =   0
         TabIndex        =   30
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
         Caption         =   "IdenLote         Fecha I.  Vencim.    Ancho     Largo        Cant   Res."
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
      BackColor       =   &H00004080&
      Height          =   4560
      Left            =   8400
      ScaleHeight     =   4500
      ScaleWidth      =   1320
      TabIndex        =   1
      Top             =   -20
      Width           =   1380
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
         Picture         =   "ACOSLOTE.frx":030A
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
         Top             =   3750
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
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "ACOSLOTE.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Salir"
         Top             =   300
         Width           =   690
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "ACOSLOTE.frx":075E
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   4100
         Width           =   1515
      End
   End
   Begin VB.Label Label31 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción:"
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
      Left            =   2520
      TabIndex        =   12
      Top             =   165
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
      Left            =   2730
      TabIndex        =   11
      Top             =   420
      Width           =   5535
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
      Left            =   420
      TabIndex        =   10
      Top             =   420
      Width           =   1860
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Código:"
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
      Left            =   -630
      TabIndex        =   9
      Top             =   165
      Width           =   1545
   End
End
Attribute VB_Name = "ACOSLOTE"
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



Private Sub Command15_Click()
   Call SELECHO("TAMONED")
   VALONUE$ = VALACT1$("TAMONED")
   If VALONUE$ <> "" Then
      Text13.TEXT = VALONUE$
   End If
End Sub

Private Sub Command2_Click()
   '\\\OT-08-0304-FL-27-05-08///
   Command2.Enabled = False
   '
   'VALIDACIONES
  
   FEFA% = FECHANUM(Label36)
   FEVE% = FECHANUM(Label38)
   CIXI% = CODINT%(Label28)
   IDLOTX$ = TRIM$(Label2)
   VALREP# = XVALO(Text1)
   MONEREP% = XVALO(Text13)
   
   '
   If CIXI% < 1 Then
      Call MENSERR(24, "Codigo no Válido")
      GoTo 99
   End If
   '
   If IDLOTX$ = "" Then
      Call MENSERR(24, "Falta Seleccionar Lote a Actualizar Costo.")
      GoTo 99
   End If
   '
   If VALREP < 0.005 Then
      Call MENSERR(24, "Costo de Reposición del Lote no Valido.")
      GoTo 99
   End If
   '
   If ECOARCH$("TAMONED", Chr$(MONEREP%)) = "" Then
      Call MENSERR(24, "Moneda de Reposición no Valida.")
      GoTo 99
   End If
   '
   If FECHANUM(Text25) = 0 Then
        Call MENSERR(24, "Fecha de Actualización de Costos Invalida")
        GoTo 99
   End If
   '
   FECREPO = CVDAT(FECHANUM(Text25))
   Call APERBASDAT("STOCKS")
   SQLX$ = " SELECT * FROM LOTINGRE "
   SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CIXI% & " "
   SQLX$ = SQLX$ + "AND IdenLote LIKE '" & IDLOTX$ & "%'"
   '
   Dim LOTINGTEMP As ADODB.Recordset
   Set LOTINGTEMP = New ADODB.Recordset
   LOTINGTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText

   With LOTINGTEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        On Error GoTo 0
        Do While Not .EOF
         !Costo_Rep = VALREP#
         !Moneco_Rep = MONEREP%
         !FechCo_Rep = FECREPO
         IDENLO$ = !IdenLote
         .Update
         Call ACSALOTE(CIXI%, IDENLO$, "NOMESS")
         .MoveNext
         Loop
      End If
   End With
   LOTINGTEMP.Close
   Set LOTINGTEMP = Nothing
   Call ARMALILOTES(CIXI%)

99 Screen.MousePointer = 1
   '
   Command2.Enabled = True
   '\\\OT-08-0304-FL-FIN///
End Sub

Private Sub Command3_Click()
   'FORHLPDF.Pdf1.LOADFILE "C:\Mis Documentos\contab04.pdf"
   Command3.Enabled = False
   CUDI$ = UCase$(CurDir)
   PPXX% = InStr(CUDI$, "\WINDOWS")
   If PPXX% > 0 Then
      rutax$ = Left$(CUDI$, PPXX%) + "MANUALES\"
      If EXISTE%(rutax$ + "TRAZAB04.PDF") > 0 Then
'         FORHLPDF.Pdf1.setShowToolbar (True)
'         On Error Resume Next
'         FORHLPDF.Pdf1.LOADFILE rutax$ + "TRAZAB04.pdf"
'         If Err Then GoTo 20
'         On Error GoTo 0
'         FORHLPDF.Show 1
'         GoTo 99
      End If
   End If
   '
20 Call MENSERR(24, "Imposible Abrir Archivo de Ayuda")
   '
99 Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Call SELECHO("SELFECHA")
    VALONUE$ = VALACT1$("SELFECHA")
    If VALONUE$ <> "" Then
      Text25.TEXT = VALONUE$
    End If
End Sub



Private Sub Command5_Click()
    Call SELECHO("MASTER")
    If VALACT1$("MASTER") <> "" Then
       Label28 = VALACT1$("MASTER")
    End If
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





































Private Sub Label24_Click()

    Label17.Caption = Label24.Caption
    
End Sub



Private Sub Label28_Change()
    CIXI% = CODINT%(Label28)
    Label30 = DESCRIT0$(CIXI%)
    Call ARMALILOTES(CIXI%)
    Call LIMPLAB
    Call LIMPTEXT
End Sub

Sub ARMALILOTES(CIXI%)
    '
     FLEXCONN.Execute (FILTSQL$("SET DATEFORMAT dmy"))
    '
    Call APERBASDAT("STOCKS")
    SQLX$ = " SELECT * FROM LOTINGRE WITH(NOLOCK)"
'    SQLX$ = SQLX$ + " INNER JOIN BOLRECEP"
'    SQLX$ = SQLX$ + " ON LOTINGRE.IDENLOTE = BOLRECEP.IDENLOTE"
    SQLX$ = SQLX$ + " WHERE LOTINGRE.Cod_Inte = " & CIXI%
'    SQLX$ = SQLX$ + " AND (CanAlta - CanConsu) > 0"    ' no trae los consumidos
    SQLX$ = SQLX$ + " ORDER BY FechAlta ASC, LOTINGRE.IdenLote ASC "
    '
    Dim LOTINGTEMP As ADODB.Recordset
    Set LOTINGTEMP = New ADODB.Recordset
    Set LOTINGTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
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
           FEX = CVINT(!FECHALTA)
         Call REPLA(TTXX$, IDLOT$, 1, 16)
         Call REPLA(TTXX$, FECHATEX$(FEX), 17, 8)
           FEX = CVINT(!FechVenc)
         Call REPLA(TTXX$, FECHATEX$(FEX), 27, 8)
           ANCHX = !AnchoMili
         Call REPLA(TTXX$, FORMATNUM$(ANCHX, "F10.1"), 35, 10)
           CSALFI# = !CanSaldo: If CSALFI# < 0 Then CSALFI# = 0
           METSAL = 0: If ANCHX > 0.1 Then METSAL = 1000 * CSALFI# / ANCHX
         Call REPLA(TTXX$, FORMATNUM$(METSAL, "F10.0"), 45, 10)
         Call REPLA(TTXX$, FORMATNUM$(CSALFI#, "F12,1"), 55, 12)
           If !CanReser > 0.05 Then
             Call REPLA(TTXX$, "SI", 69, 2)
           End If
           FEX = CVINT(!FechFab)
         Call REPLA(TTXX$, FECHATEX$(FEX), 129, 8)
         Call REPLA(TTXX$, MARCALOT$(CIXI%, IDLOT$), 137, 32)
         Call REPLA(TTXX$, LOTORIG$(CIXI%, IDLOT$), 169, 24)
         Call REPLA(TTXX$, FORMATNUM$(!DEPOUBIC, "F4.0"), 193, 4)
         Call REPLA(TTXX$, SAFETEXT(!UBICFIS), 197, 32)
         '
         Call REPLA(TTXX$, FORMATNUM$(XVALO(!Costo_Rep), "F12.4"), 229, 12)
         Call REPLA(TTXX$, FORMATNUM$(XVALO(!Moneco_Rep), "I4"), 245, 4)
         Call REPLA(TTXX$, FECHATEX$(CVINT(!FechCo_Rep)), 249, 8)
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

Private Sub List1_Click()
    
    Call LIMPTEXT
    '
    TTXX$ = List1.TEXT
    '
    NCR$ = Mid(TTXX$, 1, 12)
    COD$ = Label28
    CI% = CODINT%(COD$)
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
    '
    Text1 = TRIM$(Mid$(TTXX$, 229, 12))
    Text13 = TRIM$(Mid$(TTXX$, 245, 4))
    Text25 = TRIM$(Mid$(TTXX$, 249, 8))
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
    Label43 = ""
  
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
    Text1 = ""
    Text13 = ""
    Text25 = ""
  
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Text13.SetFocus
End Sub

Private Sub Text1_LostFocus()
  '  Text1 = FORMATNUM(XVALO(Text1), "F12.4")
End Sub

Private Sub Text13_Change()
   VALONU = XVALO(Text13.TEXT)
   If VALONU < 0 Or VALONU > 255 Then
      Text13.TEXT = "": Exit Sub
   End If
   Label3.Caption = ECOARCH$("TAMONED", Chr$(VALONU))
   '
End Sub







Private Sub Text13_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Text25.SetFocus
End Sub
