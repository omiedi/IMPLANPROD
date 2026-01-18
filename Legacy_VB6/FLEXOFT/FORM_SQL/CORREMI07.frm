VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form CORREMI07 
   Caption         =   "Correción de Cantidades en Remitos"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   10185
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   330
      Index           =   0
      Left            =   1470
      ScaleHeight     =   270
      ScaleWidth      =   1110
      TabIndex        =   45
      Top             =   525
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame8 
      Caption         =   "Agregar Registro"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      TabIndex        =   34
      Top             =   6240
      Width           =   8820
      Begin VB.TextBox Text4 
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
         Left            =   4680
         TabIndex        =   43
         Top             =   360
         Width           =   780
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
         Height          =   285
         Left            =   3240
         TabIndex        =   42
         Top             =   360
         Width           =   175
      End
      Begin VB.TextBox Text3 
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
         Left            =   1440
         TabIndex        =   41
         Top             =   360
         Width           =   1860
      End
      Begin VB.TextBox Text2 
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
         Left            =   7080
         TabIndex        =   36
         Top             =   360
         Width           =   1380
      End
      Begin VB.CommandButton Command3 
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
         Height          =   255
         Left            =   7080
         TabIndex        =   35
         Top             =   840
         Width           =   1380
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pedido:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2520
         TabIndex        =   44
         Top             =   360
         Width           =   2100
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   600
         TabIndex        =   40
         Top             =   360
         Width           =   780
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Descripcion:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   0
         TabIndex        =   39
         Top             =   840
         Width           =   1380
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
         Left            =   1440
         TabIndex        =   38
         Top             =   840
         Width           =   5340
      End
      Begin VB.Label Label19 
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
         Height          =   240
         Left            =   4920
         TabIndex        =   37
         Top             =   360
         Width           =   2100
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Frame1"
      Height          =   2775
      Left            =   7680
      TabIndex        =   31
      Top             =   1590
      Width           =   15
   End
   Begin VB.Frame Frame6 
      Caption         =   "Nuevo Valor"
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
      TabIndex        =   18
      Top             =   4440
      Width           =   8820
      Begin VB.CommandButton Command2 
         Caption         =   "Registrar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         TabIndex        =   29
         ToolTipText     =   "Corrige la Cantidad Remitida "
         Top             =   1260
         Width           =   3300
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
         Left            =   7080
         TabIndex        =   28
         Top             =   840
         Width           =   1380
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Factura:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   0
         TabIndex        =   48
         Top             =   1200
         Width           =   1380
      End
      Begin VB.Label Label18 
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
         Left            =   1440
         TabIndex        =   47
         Top             =   1200
         Width           =   3420
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad Corregida:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4920
         TabIndex        =   27
         Top             =   840
         Width           =   1980
      End
      Begin VB.Label Label14 
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
         Left            =   7080
         TabIndex        =   26
         Top             =   360
         Width           =   1380
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad Original:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4800
         TabIndex        =   25
         Top             =   360
         Width           =   2100
      End
      Begin VB.Label Label12 
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
         Left            =   1440
         TabIndex        =   24
         Top             =   840
         Width           =   3420
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Descripcion:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   0
         TabIndex        =   23
         Top             =   840
         Width           =   1380
      End
      Begin VB.Label Label10 
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
         Left            =   3000
         TabIndex        =   22
         Top             =   360
         Width           =   1860
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2160
         TabIndex        =   21
         Top             =   360
         Width           =   780
      End
      Begin VB.Label Label8 
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
         Left            =   1080
         TabIndex        =   20
         Top             =   360
         Width           =   900
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Item:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -840
         TabIndex        =   19
         Top             =   360
         Width           =   1860
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Frame1"
      Height          =   2775
      Left            =   6960
      TabIndex        =   17
      Top             =   1590
      Width           =   15
   End
   Begin VB.Frame Frame4 
      Caption         =   "Frame1"
      Height          =   2775
      Left            =   6120
      TabIndex        =   16
      Top             =   1590
      Width           =   15
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame1"
      Height          =   2775
      Left            =   5280
      TabIndex        =   15
      Top             =   1590
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame1"
      Height          =   2775
      Left            =   2360
      TabIndex        =   14
      Top             =   1590
      Width           =   15
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2775
      Left            =   680
      TabIndex        =   13
      Top             =   1590
      Width           =   15
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      ItemData        =   "CORREMI07.frx":0000
      Left            =   240
      List            =   "CORREMI07.frx":0002
      TabIndex        =   11
      Top             =   1590
      Width           =   8820
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00004040&
      Height          =   8040
      Left            =   9240
      ScaleHeight     =   7980
      ScaleWidth      =   960
      TabIndex        =   3
      Top             =   0
      Width           =   1020
      Begin VB.CommandButton Command8 
         Caption         =   "Kdx."
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
         Left            =   140
         Picture         =   "CORREMI07.frx":0004
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Ficha de Movimientos de Material"
         Top             =   1995
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
         Left            =   140
         Picture         =   "CORREMI07.frx":030E
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   960
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   5
         Top             =   5280
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   6
            Top             =   0
            Width           =   12000
         End
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
         Left            =   140
         Picture         =   "CORREMI07.frx":0618
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   210
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -270
         Picture         =   "CORREMI07.frx":0762
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   5640
         Width           =   1515
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
      Left            =   8880
      TabIndex        =   0
      Top             =   255
      Width           =   175
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "CORREMI07.frx":2684
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "CORREMI07.frx":28B8
      TabIndex        =   46
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label45 
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
      Left            =   8600
      TabIndex        =   51
      Top             =   600
      Width           =   420
   End
   Begin VB.Label Label22 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Mes"
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
      Left            =   6645
      TabIndex        =   50
      Top             =   615
      Width           =   405
   End
   Begin VB.Label Label44 
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
      Left            =   7050
      TabIndex        =   49
      Top             =   600
      Width           =   1485
   End
   Begin VB.Label Label16 
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
      Left            =   1890
      TabIndex        =   32
      Top             =   255
      Width           =   660
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Item  Codigo                  Descripcion                                     Entr.        Dev.      Ped.          P.Unit."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   1350
      Width           =   8820
   End
   Begin VB.Label Label4 
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
      Left            =   8130
      TabIndex        =   10
      Top             =   990
      Width           =   900
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7320
      TabIndex        =   9
      Top             =   1005
      Width           =   660
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
      Left            =   1905
      TabIndex        =   8
      Top             =   750
      Width           =   4515
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Numero de Cliente:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   0
      TabIndex        =   7
      Top             =   300
      Width           =   1860
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
      Left            =   7050
      TabIndex        =   2
      Top             =   255
      Width           =   1860
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Numero de Remito:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5160
      TabIndex        =   1
      Top             =   300
      Width           =   1860
   End
End
Attribute VB_Name = "CORREMI07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub CorrCantiFacturar()
    DIFE = XVALO(Label14) - XVALO(Text1)
    If DIFE <= 0.05 Then
      Call COMUNI("Atención!!! \Opción de Modificar Cantidad Para Facturación Parcial\Solo Por Cantidades Menores a la Activa")
      Call PINTATEXT(Text1)
      Exit Sub
    End If
    '
    CIXI% = CODINT%(Label10)
    If CIXI% > 0 Then
    '
    
        NORIT% = XVALO(Label8)
        CMO$ = TRIM$(Control$("", "CODREFAC"))
        If CMO$ = "" Then CMO$ = "RT"
        '
        'ACTUALIZA EL REGISTRO ORIGINAL POR LA DIFERENCIA
        'CONDICION PARA EL REGISTRO UNICO
        CONDIDIFE$ = "codiMovi = '" & CMO$ & "' "
        CONDIDIFE$ = CONDIDIFE$ + " AND DoPriCor = '" & Label28.Caption & "' "
        CONDIDIFE$ = CONDIDIFE$ + " AND MONTH(FECHMOV) = " & XVALO(Label45.Caption)
        CONDIDIFE$ = CONDIDIFE$ + " AND Cod_Inte = " & CIXI%
        CONDIDIFE$ = CONDIDIFE$ + " AND NuOrItem = " & NORIT%
        '
        On Error GoTo ERROR_GUARDAR
        FLEXCONN.BeginTrans
        Call ACTUREGISTRO("MOVISTO", "CANTSALID", CONDIDIFE$, DIFE, REGAF&, "NOMESS") 'ACTUALIZA UN SOLO REGISTRO
        '
        'CONDICION PARA OBTENER EL MAXIMO DEL OREDEN DE TODOS LOS ITEMS DEL REMITO  ACTIVO
        CONDI$ = "codiMovi = '" & CMO$ & "' "
        CONDI$ = CONDI$ + " AND DoPriCor = '" & Label28.Caption & "' "
        CONDI$ = CONDI$ + " AND MONTH(FECHMOV) = " & XVALO(Label45.Caption)

        'TOMO EL VALOR MAXIMO DE ORDEN + 1
        MAXNUOR% = 1 + XVALO(VALOBADA("MOVISTO", "MAX(NuOrItem)", CONDI$, "MOMESS"))
        
        'GENERAR UN NUEVO REGISTRO POR LA CANTIDAD INGRESADA

        NUEVALOR = Abs(XVALO(Text1))
        SQLX$ = " INSERT INTO MOVISTO  (CantIngre,CantSaldo,CantSalid,Causa,Cod_Exte,Cod_Inte,CodCarpImp,CodiEmpr,CodiMovi,"
        SQLX$ = SQLX$ + " CostoRep,DeLaRemi,DepoMovi,Descri_Lar,Descrip,DocuOrig,DoPriCor,DoPriLar,DOPRINUM,DoSecCor,DoSecLar,"
        SQLX$ = SQLX$ + " FechMov,FeReMovi,FHor_Verif,IdCalidad,IdenLote,IdenMovi,LotConsum,MoneVal,NroEnvio,Nume_Clie,"
        SQLX$ = SQLX$ + " Nume_Prov,NumeOrdPed,NUMPVTA,NumRem_Cli,NumUsuar,NuorItem,NuPedCli,OCOMPRA_RTOABIERTO,"
        SQLX$ = SQLX$ + " PLC_Has,PNL_Des,ReferCor,Sucu_Clie,TipoCam,TipoTran,User_Tran,ValoUnit,ValoUnitIva,ValTotIt,ValTotItIva,"
        SQLX$ = SQLX$ + " Verificado,VerificadoPor )"
        'INSERTA UN NUEVO REGISTRO CON 3 CAMBIOS LA CANTIDAD (NUEVALOR) Y EL NUMERO DE ORDEN (MAXNUOR%) Y EL CODIGO DE MOVIMIENTO (CMO$)
        SQLX$ = SQLX$ + " SELECT CantIngre , CantSaldo," & NUEVALOR
        SQLX$ = SQLX$ + " , Causa, Cod_Exte, Cod_Inte, CodCarpImp, CodiEmpr,'" & CMO$ & "', "
        SQLX$ = SQLX$ + " CostoRep,DeLaRemi,DepoMovi,Descri_Lar,Descrip,DocuOrig,DoPriCor,DoPriLar,DOPRINUM,DoSecCor,DoSecLar,"
        SQLX$ = SQLX$ + " FechMov,FeReMovi,FHor_Verif,IdCalidad,IdenLote,IdenMovi,LotConsum,MoneVal,NroEnvio,Nume_Clie,"
        SQLX$ = SQLX$ + " Nume_Prov,NumeOrdPed,NUMPVTA,NumRem_Cli,NumUsuar," & MAXNUOR%
        SQLX$ = SQLX$ + " ,NuPedCli,OCOMPRA_RTOABIERTO,"
        SQLX$ = SQLX$ + " PLC_Has,PNL_Des,ReferCor,Sucu_Clie,TipoCam,TipoTran,User_Tran,ValoUnit,ValoUnitIva,ValTotIt,ValTotItIva,"
        SQLX$ = SQLX$ + " Verificado,VerificadoPor FROM MOVISTO"
        
        SQLX$ = SQLX$ + " WHERE " & CONDIDIFE$
        FLEXCONN.Execute (SQLX$)
        
ERROR_GUARDAR:
        If Err.Number <> 0 Then
           FLEXCONN.RollbackTrans
           Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description)
           On Error GoTo 0
           GoTo 49
           Exit Sub
        Else
          FLEXCONN.CommitTrans
          Screen.MousePointer = 1
          Call COMUNI("Correccion Procesada.")
          Call Label28_Change
          Label8.Caption = ""
          Label10.Caption = ""
          Label14.Caption = ""
          Label12.Caption = ""
          Text1.TEXT = ""
          Exit Sub
        End If
    End If
        
  
'        'aqui PROCESAR LOS KITS
'        If DESCOMPO%(CIXI%) > 0 Then
'           FF% = CVINT(VALOBADA("MOVISTO", "FechMov", CONDI$))
'           LOTE$ = (VALOBADA("MOVISTO", "IdenLote", CONDI$))
'           NROREMIT$ = (VALOBADA("MOVISTO", "DoPriLar", CONDI$))
'           REMICOR1$ = Label28.Caption
'           DOSELA$ = (VALOBADA("MOVISTO", "DoSecLar", CONDI$))
'           DOSEC$ = (VALOBADA("MOVISTO", "DoSecCor", CONDI$))
'           REFEMO$ = (VALOBADA("MOVISTO", "ReferCor", CONDI$))
'           NUORIT% = XVALO(VALOBADA("MOVISTO", "MAX(NuOrItem)", CONDIBA$)) + 1
'           VUNIST# = XVALO((VALOBADA("MOVISTO", "ValoUnit", CONDI$)))
'           MONEVA% = XVALO((VALOBADA("MOVISTO", "MoneVal", CONDI$)))
'           NC% = XVALO((VALOBADA("MOVISTO", "Nume_Clie", CONDI$)))
'           DEPO% = XVALO((VALOBADA("MOVISTO", "DepoMovi", CONDI$)))
'           '
'           Call MOVISTOK(FF%, CIXI%, LOTE$, "CI", NROREMIT$, REMICOR1$, DOSELA$, DOSEC$, REFEMO$, NUORIT%, VUNIST#, MONEVA%, NC%, DEPO%, DIFE)
'           '
'           'MOVIMIENTO SALIDA PARA COMPONENTES DEL KIT
'           Call AREXPLO(CIXI%, 1, 1)
'           For KKJ% = 1 To CAIT%
'              NUORIT% = NUORIT% + 1
'              CICO% = CIJ%(KKJ%)
'              CACICO = USOI(KKJ%) * DIFE * (-1)
'              Call MOVISTOK(FF%, CICO%, LOTE$, "CI", NROREMIT$, REMICOR1$, DOSELA$, DOSEC$, REFEMO$, NUORIT%, VUNIST#, MONEVA%, NC%, DEPO%, CACICO)
'           Next KKJ%
'        End If
        '
'        Screen.MousePointer = 1
'        Call COMUNI("Correccion Procesada.")
'        Call Label28_Change
        '
'      End If
    '
49  Screen.MousePointer = 1
    Label8.Caption = ""
    Label10.Caption = ""
    Label14.Caption = ""
    Label12.Caption = ""
    Text1.TEXT = ""

End Sub

Private Sub Command1_Click()
    '
    Label28.Caption = ""
    Call LIMCORREMI
    Call CIERRARCH("*.*")
    Hide
    '
End Sub

Private Sub Command2_Click()
    '
    If Left$(TRIM$(Label18), 3) = "FC-" Or Left$(TRIM$(Label18), 3) = "NC-" Or Left$(TRIM$(Label18), 3) = "ND-" Then
        Call MENSERR(24, "Imposible Corregir Cantidad\Registro Seleccionado Yá Facturado")
        Exit Sub
    End If

    If CODINT%(TRIM(Label10.Caption)) = 0 Then
        Call MENSERR(24, "Codigo Invalido"): Exit Sub
    End If
    '
    If TRIM$(Text1) <> "" Then
      If IsNumeric(TRIM$(Text1.TEXT)) = False Then
        Call MENSERR(24, "Cantidad Corregida Invalida, Solo Valores Numéricos")
        Exit Sub
      End If
    End If
    '
    If XVALO(Text1) * XVALO(Label14) < 0 Then
        Call MENSERR(24, "Debe Coincidir el Signo\de la Cantidad Original\y la Cantidad Corregida")
        Exit Sub
    End If
    '
    
    DIFE = XVALO(Text1) - XVALO(Label14)
    '
    If Command2.Caption = "Guardar Cantidad Parcial" Then
       Call CorrCantiFacturar 'modalidad usada por Fuinyter para de un remito facturar cantidades parciales
       'esto va a permitir Ej: remitir 20 u. y facturar 15 y dejar pendiente 5 , el sistema crea un registro nuevo
       'por 5 u. y corrige el original de 20 modificandolo a 15  u.
       Exit Sub
    End If
    
    If Abs(DIFE) >= 0.05 Then
      CIXI% = CODINT%(Label10)
      If CIXI% > 0 Then
        CMO$ = TRIM$(Control$("", "CODREFAC"))
        If CMO$ = "" Then CMO$ = "RT"
        '
        NORIT% = XVALO(Label8)
        NUEVALOR = Abs(XVALO(Text1))
        Call APERBASDAT("STOCKS")
        CONDI$ = "(CodiMovi = '" & CMO$ & "' OR CodiMovi = 'DV') "
        CONDI$ = CONDI$ + " AND DoPriCor = '" & Label28.Caption & "' "
           CONDIBA$ = CONDI$   ' ESTA SE USA PARA EL NUMERO DE ORDEN DEL ITEM
        CONDI$ = CONDI$ + " AND Cod_Inte = " & CIXI%
        CONDI$ = CONDI$ + "AND NuOrItem = " & NORIT%
        '
        Call ACTUREGISTRO("MOVISTO", "CantIngre", CONDI$, 0, REGAFE&)
        Call ACTUREGISTRO("MOVISTO", "CantSalid", CONDI$, 0, REGAFE&)
        CAMPO$ = "Cantsalid": If Sgn(XVALO(Text1)) < 0 Then CAMPO$ = "CantIngre"
        Call ACTUREGISTRO("MOVISTO", CAMPO$, CONDI$, NUEVALOR, REGAFE&)
        CODMO$ = CMO$: If Sgn(XVALO(Text1)) < 0 Then CODMO$ = "DV"
        Call ACTUREGISTRO("MOVISTO", "CodiMovi", CONDI$, CODMO$, REGAFE&)
        '
        'aqui PROCESAR LOS KITS
        If DESCOMPO%(CIXI%) > 0 Then
           FF% = CVINT(VALOBADA("MOVISTO", "FechMov", CONDI$))
           LOTE$ = (VALOBADA("MOVISTO", "IdenLote", CONDI$))
           NROREMIT$ = (VALOBADA("MOVISTO", "DoPriLar", CONDI$))
           REMICOR1$ = Label28.Caption
           DOSELA$ = (VALOBADA("MOVISTO", "DoSecLar", CONDI$))
           DOSEC$ = (VALOBADA("MOVISTO", "DoSecCor", CONDI$))
           REFEMO$ = (VALOBADA("MOVISTO", "ReferCor", CONDI$))
           NUORIT% = XVALO(VALOBADA("MOVISTO", "MAX(NuOrItem)", CONDIBA$)) + 1
           VUNIST# = XVALO((VALOBADA("MOVISTO", "ValoUnit", CONDI$)))
           MONEVA% = XVALO((VALOBADA("MOVISTO", "MoneVal", CONDI$)))
           NC% = XVALO((VALOBADA("MOVISTO", "Nume_Clie", CONDI$)))
           DEPO% = XVALO((VALOBADA("MOVISTO", "DepoMovi", CONDI$)))
           '
           Call MOVISTOK(FF%, CIXI%, LOTE$, "CI", NROREMIT$, REMICOR1$, DOSELA$, DOSEC$, REFEMO$, NUORIT%, VUNIST#, MONEVA%, NC%, DEPO%, DIFE)
           '
           'MOVIMIENTO SALIDA PARA COMPONENTES DEL KIT
           Call AREXPLO(CIXI%, 1, 1)
           For KKJ% = 1 To CAIT%
              NUORIT% = NUORIT% + 1
              CICO% = CIJ%(KKJ%)
              CACICO = USOI(KKJ%) * DIFE * (-1)
              Call MOVISTOK(FF%, CICO%, LOTE$, "CI", NROREMIT$, REMICOR1$, DOSELA$, DOSEC$, REFEMO$, NUORIT%, VUNIST#, MONEVA%, NC%, DEPO%, CACICO)
           Next KKJ%
        End If
        '
        Screen.MousePointer = 1
        Call COMUNI("Correccion Procesada.")
        Call Label28_Change
        GoTo 49
        '
      End If
    End If
    '
49  Screen.MousePointer = 1
    Label8.Caption = ""
    Label10.Caption = ""
    Label14.Caption = ""
    Label12.Caption = ""
    Text1.TEXT = ""
    
'    OPX% = COMALTER%("Desea Re-Imprimir el Documento\\No, No Re-Imprimir\Si, Imprimir")
'    If OPX% = 2 Then
'        PRINTERPORT$ = TRIM$(CONTROL$(IDENTER$, "PORTREM"))
'     'MsgBox "printerport = *" + PRINTERPORT$ + "*"
'        FORIPRE$ = TRIM$(CONTROL$("", "FOREMIFA"))
'
'     'MsgBox "foripre = *" + FORIPRE$ + "*"
'        If TRIM$(FORIPRE$) <> "" Then
'          '
'    '      PREFIREM$ = TRIM(Label10)
'    '      If PREFIREM$ = "" Then PREFIREM$ = "0001"
'    '      NROREMIT$ = TRIM$(Mid$(REMI$, 4))
'    '      While Len(NROREMIT$) < 8: NROREMIT$ = "0" + NROREMIT$: Wend
'    '      NROREMIT$ = PREFIREM$ + "-" + NROREMIT$
'          '
'          Call LIMPIATETAB
'          FEX = FF%
'          FECHDOC$ = FECHATEX$(FEX)
'          NUMEDOC$ = NROREMIT$
'          TIPODOC$ = "Remito"
'          If Option2.Value = True Then
'            TIPODOC$ = "Remito (en Consignación)"
'          End If
'          '
'          CODPARAM$(1) = "TRANSPOR": DOCPARAM$(1) = ""
'          CODPARAM$(2) = "NUBULTOS": DOCPARAM$(2) = ""
'          CODPARAM$(3) = "CUIT-TRA": DOCPARAM$(3) = ""
'          CODPARAM$(4) = "VALO-DEC": DOCPARAM$(4) = ""
'          CODPARAM$(5) = "DOMI-TRA": DOCPARAM$(5) = ""
'          CODPARAM$(6) = "DOCUCOND": DOCPARAM$(6) = ""
'          CODPARAM$(7) = "DEPOSAL ": DOCPARAM$(7) = ""
'          'CODPARAM$(8) = "NRO-REMI": DOCPARAM$(8) = TRIM$(NROREMIT$)
'          CODPARAM$(8) = "NRO-REMI": DOCPARAM$(8) = "" ' SE IMPRIME FORMATO 0000-00000000
'          CODPARAM$(9) = "NRO-OCOM": DOCPARAM$(9) = TRIM$(Text20.TEXT)
'          CODPARAM$(10) = "PATE-TRA": DOCPARAM$(10) = TRIM$(Text18.TEXT)
'          CODPARAM$(11) = "N-AUTENT": DOCPARAM$(11) = TRIM$(NORVENTA)
'          CODPARAM$(12) = "COND-TRA": DOCPARAM$(12) = TRIM$(Text19)
'          CODPARAM$(13) = "LOCA-TRA": DOCPARAM$(13) = TRIM$(Text21)
'          CODPARAM$(14) = "NRO-AENT": DOCPARAM$(14) = TRIM$(Text22)
'
'          '\\\OT-07-0007-RG-08/01/07///
'          HORA$ = HORAENT$(NCLI)
'          CODPARAM$(15) = "Z-PLANO": DOCPARAM$(15) = TRIM$(HORA$)
'          CODPARAM$(16) = "IMP-TOTA": DOCPARAM$(16) = FORMATNUM$(ACUMLAIMPOX#, "F14.2")
'          CODPARAM$(17) = "DOCU-APL": DOCPARAM$(17) = TRIM$(STRNPEDI$)
'          CODPARAM$(18) = "DESTINO": DOCPARAM$(18) = TRIM$(RASOCLI$(NCLI)) + " - Almacén " + Label11
'
'          '\\\OT-07-0007-RG-FIN///
'          CAPARDOC% = 18
'          '
'          CODTABU1$(1) = "COD-MAT"
'          CODTABU1$(2) = "DES-MAT"
'          CODTABU1$(3) = "CANTIDAD"
'          CODTABU1$(4) = "NRO-LOTE"
'          CODTABU1$(5) = "UNIMED"
'          CODTABU1$(6) = "N-OCOMPR"
'          CODTABU1$(7) = "ORD-ITEM"
'          CODTABU1$(8) = "REF-MAT1"
'          CODTABU1$(9) = "REF-MAT2"
'          CODTABU1$(10) = "PRE-NETO"
'          CODTABU1$(11) = "NET-ITEM"
'          CODTABU1$(12) = "REF-MAT6"
'          CODTABU1$(13) = "REF-MAT5"
'          CACOLTAB1% = 13
'          '
'          DESTIDOCNUE = NCLI
'          CAITAB1% = 0
'          '
'          For I = 0 To List1.ListCount - 1
'           XX$ = List1.LIST(I)
'           LOTE$ = TRIM$(Mid$(XX$, 85))
'
'          For YX& = 1 To List1.ListCount - 1
'            YY$ = List1.LIST(I)
'            CI% = XVALO(Mid$(YY$, 1, 15))
'            CAN# = CVD(Mid$(YY$, 89, 8))
'            CAC# = 0
'            If CI% > 0 Or Frame15.Visible = True Then
'              If CI% <= NUMAS% Or Frame15.Visible = True Then
'                If CAN# + CAC# > 0.005 Or Frame15.Visible = True Then
'                  '
'                  UNID$ = UNIMED$(CI%)
'                  If UNID$ = "" Then UNID$ = UNIMED$(CI%)
'                  CANS$ = CSTRING$(MKD$(CAN#), 3, 12, 1, 2)
'                  If CAC# >= 0.05 Then
'                    CANS$ = CSTRING$(MKD$(CAC#), 3, 12, 1, 2)
'                    Call REPLA(CANS$, "(*)", 1, 3)
'                  End If
'                  NULOTE$ = TRIM$(Mid$(YY$, 107, 12))
'                  NUOCO$ = CSTRING$(Mid$(YY$, 151, 4), 3, 7, 0, 2)
'                  ORD& = XVALO(Mid$(YY$, 241, 5)) ' NUMERO DE ORDEN INGRESADO EN VENTANA DE ARCON NEUQUEN.
'                  '
'                  SERINU$ = TRIM$(Mid$(YY$, 209, 16)) + TRIM$(Mid$(YY$, 225, 8))
'                  If SERINU$ <> "" Then
'                    If TRIM$(Mid$(YY$, 233, 8)) <> "" Then
'                      SERINU$ = SERINU$ + " - "
'                    End If
'                  End If
'                  SERINU$ = SERINU$ + TRIM$(Mid$(YY$, 233, 8))
'                  '
'                  CAITAB1% = CAITAB1% + 1
'                  CODDX$ = CODEXT$(CI%)
'                  If TRIM$(Mid$(YY$, 97, 8)) <> "" Then 'VALIDACION POR QUE SI NO QUEDA EL GUION SOLO
'                    CODDX$ = CODEXT$(CI%) + "-" + TRIM$(Mid$(YY$, 97, 8)) ' agrega subcodigo
'                  End If
'                  If REPLAORI% = 1 Then
'                    CODDX$ = CODORIG$(CI%, NCLI)
'                  End If
'                  '
'                  TETABU1$(1, CAITAB1%) = CODDX$
'                  TETABU1$(2, CAITAB1%) = DESCRIT0$(CI%)
'                  If Frame15.Visible = True Then
'                    TETABU1$(2, CAITAB1%) = Mid$(YY$, 5, 64)
'                  End If
'                  '
'                  TETABU1$(3, CAITAB1%) = CANS$
'                  TETABU1$(4, CAITAB1%) = NULOTE$
'                  TETABU1$(5, CAITAB1%) = UNID$
'                  TETABU1$(6, CAITAB1%) = NUOCO$
'                  TETABU1$(7, CAITAB1%) = Str$(CAITAB1%)
'                  TETABU1$(8, CAITAB1%) = NUOCOCLI$(Val(NUOCO$), CI%)
'                  TETABU1$(9, CAITAB1%) = NUPLANO$(CI%) '
'                  '
'                  'VALORIZACION ALEJPA SEGUN LISTA VIGENTE
'                  IMPOUNI# = 0: IMPORTEX# = 0
'                  If InStr(UCase$(EMPREAC$), "ALEJPA") > 0 Then
'                    IMPOUNI# = PRELISTACT#(CI%, TIMONE%)
'                    IMPOUNI# = IMPOUNI# * TICAMONE(TIMONE%) 'VALOR UNITARIO EN PESOS
'                    IMPORTEX# = CAN# * IMPOUNI#             'IMPORTE EN PESOS
'                    ACUMLAIMPOX# = ACUMLAIMPOX# + IMPORTEX# 'IMPORTE ACUMULADO EN PESOS
'                  End If
'                  TETABU1$(10, CAITAB1%) = FORMATNUM$(IMPOUNI#, "F12.2")
'                  TETABU1$(11, CAITAB1%) = FORMATNUM$(IMPORTEX#, "F12.2")
'                    CORIG$ = TRIM$(CODORIG$(CI%, NCLI))
'                    'If CORIG$ = TRIM$(CODDX$) Then CORIG$ = ""
'                    If CORIG$ = CODEXT$(CI%) Then CORIG$ = ""
'                    'If CORIG$ <> CODEXT$(CI%) Then CORIG$ = CODEXT$(CI%) 'ESTA LINEA HABRIA QUE QUITAR
'                  TETABU1$(12, CAITAB1%) = CORIG$
'                  'FIN ALEJPA
'                  '
'                  TETABU1$(13, CAITAB1%) = SAFETEXT$(ORD&) ' NUMERO DE ORDEN INGRESADO EN VENTANA DE ARCON NEUQUEN.
'                  'IMPRESION DE ESPECIFICACION
'                  If CONTROL("FACTUREM", "INCLUSPE") = "SI" Then ' IMPRIME ESPECIFICACIONES DEL ARTICULO
'                    XX$ = ESPECIFIT$(CI%)
'                    Call PRINTEXT(XX$, 2, "DES-MAT", FORIPRE$)
'                  End If
'                  Call CARLINSEP         'Carga linea de separacion entre codigo y codigo
'                  '
'                  If SERINU$ <> "" Then
'                    CAITAB1% = CAITAB1% + 1
'                    For KI1% = 1 To 20
'                      TETABU1$(KI1%, CAITAB1%) = ""
'                    Next KI1%
'                    TETABU1$(2, CAITAB1%) = "S.No." + SERINU$
'                  End If
'                  '
'                  Call DETLOTREM(CI%, YX&)
'                  '
'                  'NCIY% = NCLI
'                  CIIY% = CI%
'                  'Call DATADI(NCIY%, CIIY%, ORCO$, NORI$, NPLA$)
'                End If
'              End If
'            End If
'          Next YX&
'          '
'     'MsgBox "antes de PRINTDOC(FOREMIFA)"
'          Call PRINTDOC("FOREMIFA")   ' Remito de Facturacion
'          '
'        End If
'        '
'
'
'
'    End If
End Sub
Private Sub Command3_Click()
    '
    If Label28.Caption = "" Then
        Call MENSERR(24, "Debe Seleccionar el Remito"): Exit Sub
    End If
    '
    CIXI% = CODINT%(TRIM(Text3.TEXT))
    If CIXI% < 1 Then
        Call MENSERR(24, "Codigo no Valido"): Exit Sub
    End If
    '
    If XVALO(Text2.TEXT) < 0.05 Then
        Call MENSERR(24, "Cantidad no Valida"): Exit Sub
    End If
    '
    If IsNumeric(Text4.TEXT) = False Then
        Call MENSERR(24, "Número de Pedido no Valido"): Exit Sub
    End If
    '
    CMO$ = TRIM$(Control$("", "CODREFAC"))
    If CMO$ = "" Then CMO$ = "RT"
    '
    Call APERBASDAT("STOCKS")
    CONDI$ = "(CodiMovi = '" & CMO$ & "' OR CodiMovi = 'DV') "
    CONDI$ = CONDI$ + " AND DoPriCor = '" & Label28.Caption & "' "
    NORIT% = 1 + XVALO(VALOBADA("MOVISTO", "MAX(NuOrItem)", CONDI$)) 'OBTIENE NUEVO NUMERO DE ITEM
    '
    FF% = CVINT(VALOBADA("MOVISTO", "FechMov", CONDI$))
    LOTE$ = ""
    NROREMIT$ = (VALOBADA("MOVISTO", "DoPriLar", CONDI$))
    REMICOR1$ = Label28.Caption
    DOSELA$ = ""
    DOSEC$ = ""
    REFEMO$ = (VALOBADA("MOVISTO", "ReferCor", CONDI$))
    VUNIST# = 0
    MONEVA% = 1
    NC% = XVALO((VALOBADA("MOVISTO", "Nume_Clie", CONDI$)))
    DEPO% = XVALO((VALOBADA("MOVISTO", "DepoMovi", CONDI$)))
    CANTT = (-1) * XVALO(Text2.TEXT)
    NUMPEDC& = XVALO(Text4.TEXT)
    '
    CONDI$ = "(CodiMovi = '" & CMO$ & "' OR CodiMovi = 'DV' OR CodiMovi = 'CI') "
    CONDI$ = CONDI$ + " AND DoPriCor = '" & Label28.Caption & "' "
    NORIT% = 1 + XVALO(VALOBADA("MOVISTO", "MAX(NuOrItem)", CONDI$)) 'OBTIENE NUEVO NUMERO DE ITEM
    
    Call MOVISTOK(FF%, CIXI%, LOTE$, CMO$, NROREMIT$, REMICOR1$, DOSELA$, DOSEC$, REFEMO$, NORIT%, VUNIST#, MONEVA%, NC%, DEPO%, CANTT, NUMPEDC&)
    '
    'aqui PROCESAR LOS KITS
    If DESCOMPO%(CIXI%) > 0 Then
       '
       NORIT% = NORIT% + 1
       Call MOVISTOK(FF%, CIXI%, LOTE$, "CI", NROREMIT$, REMICOR1$, DOSELA$, DOSEC$, REFEMO$, NORIT%, VUNIST#, MONEVA%, NC%, DEPO%, (-1) * CANTT, NUMPEDC&)
       '
       'MOVIMIENTO SALIDA PARA COMPONENTES DEL KIT
       Call AREXPLO(CIXI%, 1, 1)
       For KKJ% = 1 To CAIT%
          NORIT% = NORIT% + 1
          CICO% = CIJ%(KKJ%)
          CACICO = USOI(KKJ%) * CANTT
          Call MOVISTOK(FF%, CICO%, LOTE$, "CI", NROREMIT$, REMICOR1$, DOSELA$, DOSEC$, REFEMO$, NORIT%, VUNIST#, MONEVA%, NC%, DEPO%, CACICO)
       Next KKJ%
    End If
    '
    Screen.MousePointer = 1
    Call COMUNI("Registro Agregado.")
    Call Label28_Change
    Exit Sub
    '
49  Screen.MousePointer = 1
    '
End Sub

Private Sub Command4_Click()
    Call SELECHO("MASTER")
    If VALACT1$("MASTER") <> "" Then
      Text3 = VALACT1$("MASTER")
    End If
End Sub

Private Sub Command5_Click()
    '
    RERE$ = REGSEL$("!SELREMI")
    If Len(RERE$) <= 4 Then Exit Sub
    Label28.Caption = TRIM$(Left$(RERE$, 10))
    
    
    '
End Sub


Private Sub Command8_Click()
    Command8.Enabled = False
    FIKARDEX09.Show 1
    Command8.Enabled = True
End Sub

Private Sub Form_Load()
   Call APLICACIONSKINS(Me, Picture2)

End Sub

Private Sub Label28_Change()
    '
    Call LIMCORREMI
    '
    List1.Clear
    CMO$ = TRIM$(Control$("", "CODREFAC"))
    If CMO$ = "" Then CMO$ = "RT"
    '
    
    If Label28 <> "" Then
        Call APERBASDAT("STOCKS")
        CONDI$ = "(CodiMovi = '" & CMO$ & "' OR CodiMovi = 'DV') "
        CONDI$ = CONDI$ + " AND DoPriCor = '" & Label28.Caption & "' "
        CONDI$ = CONDI$ + " AND MONTH(FECHMOV) = " & XVALO(Label45)
          CONDIBA$ = CONDI$
        CONDI$ = CONDI$ + " ORDER BY NuOrItem"
        Label16.Caption = VALOBADA("MOVISTO", "Nume_Clie", CONDI$)
        Label2 = rasocli$(XVALO(Label16))
        FECHH% = CVINT(VALOBADA("MOVISTO", "FechMov", CONDI$))
        Label4 = FECHATEX$(FECHH%)
        '
        Call CARGALISTA(List1, "MOVISTO", "NuOrItem/I3>;Cod_Inte/C15<;DESCRIP/A27;Cantsalid/F7.1;CantIngre/F7.1;NuPedCli/I6;ValoUnit/F10.2", CONDI$)
        Command2.Enabled = True
        '
        For i% = 1 To List1.ListCount
           AKX1$ = Left$(List1.List(i% - 1), 19)
           For J% = i% + 1 To List1.ListCount
             AKX2$ = Left$(List1.List(J% - 1), 19)
             If XVALO(Left$(AKX1$, 3)) < 1 Or AKX1$ = AKX2$ Then
                Call MENSERR(24, "Imposible Corregir este Remito.\Coincidencia de Indices en Registros " & i% & " y " & J% & ".\  \Pulse 'Aceptar' para Re-Generar los Indices\y Repita la Peticion.")
                Screen.MousePointer = 11
                Call RENUMERA("MOVISTO", "NuOrItem", CONDIBA$)
                Screen.MousePointer = 1
                Unload Me
                Exit Sub
             End If
           Next J%
        Next i%
    End If
End Sub

Private Sub List1_Click()
    '
    Static CMO$
    If CMO$ = "" Then
       CMO$ = TRIM$(Control$("", "CODREFAC"))
       If CMO$ = "" Then CMO$ = "RT"
    End If
    '
    For XX% = 0 To List1.ListCount - 1
       If List1.Selected(XX%) = True Then
          TTXX$ = List1.List(XX%)
          Label8.Caption = TRIM$(Mid$(TTXX$, 1, 3))
          Label10 = TRIM$(Mid$(TTXX$, 5, 15))      ' CODIGO EXTERNO
          GoTo 10
       End If
    Next XX%
    Exit Sub
    '
10  CIXI% = CODINT%(Label10)
    If CIXI% < 1 Then Exit Sub
    '
    CONDI$ = "(CodiMovi = '" & CMO$ & "' OR CodiMovi = 'DV') "
    CONDI$ = CONDI$ + " AND NuorItem = " & XVALO(Label8.Caption)
    CONDI$ = CONDI$ + " AND Cod_Inte = " & CIXI%
    CONDI$ = CONDI$ + " AND DoPriCor = '" & Label28.Caption & "'"
    '
    Label12.Caption = DESCRIT0$(CIXI%)
    CANSAL = XVALO(VALOBADA("MOVISTO", "CantSalid", CONDI$))
    CANDEV = XVALO(VALOBADA("MOVISTO", "CantIngre", CONDI$))
    Label14.Caption = FORMATNUM(XVALO(CANSAL - CANDEV), "F7.1")
    Text1.TEXT = FORMATNUM(XVALO(CANSAL - CANDEV), "F7.1")
    Label18 = TRIM$(VALOBADA("MOVISTO", "Doseclar", CONDI$))
    '
End Sub

Private Sub Text1_LostFocus()
    '
    Text1 = FORMATNUM$(XVALO(Text1), "F7.1")
    '
End Sub

Sub LIMCORREMI()
    '
    Label16.Caption = ""
    Label2.Caption = ""
    Label4.Caption = ""
    Label8.Caption = ""
    Label10.Caption = ""
    Label14.Caption = ""
    Label12.Caption = ""
    Text1.TEXT = ""
    List1.Clear
    Text2.TEXT = ""
    Text3.TEXT = ""
    Text4.TEXT = ""
    Label20.Caption = ""
    Label18 = ""
    '
End Sub

Private Sub Text2_LostFocus()
    Text2.TEXT = FORMATNUM$(XVALO(Text2.TEXT), "F7.1")
End Sub

Private Sub Text3_Change()
    Label20.Caption = DESCRIT$(CODINT%(Text3.TEXT))
End Sub

