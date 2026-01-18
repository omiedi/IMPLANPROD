VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form ESTASTK07 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Estadísticas de Stock por Tipo de Movimiento"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   7950
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
   ScaleHeight     =   5445
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3720
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   240
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   41
      Top             =   960
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00CDDADA&
      Caption         =   "COMPONENTES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   4755
      TabIndex        =   36
      Top             =   1785
      Width           =   2150
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Incluir Componentes Multinivel"
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
         Height          =   435
         Left            =   210
         TabIndex        =   37
         Top             =   315
         Width           =   1905
      End
   End
   Begin VB.CommandButton Command12 
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
      Left            =   6360
      Picture         =   "ESTASTK07.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Emite Reporte Solicitado"
      Top             =   4680
      Width           =   540
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00CDDADA&
      Caption         =   "TIPOS DE MOVIMIENTOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   60
      TabIndex        =   27
      Top             =   2700
      Width           =   3405
      Begin VB.ListBox List2 
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
         Height          =   2280
         Left            =   160
         Style           =   1  'Checkbox
         TabIndex        =   29
         Top             =   270
         Width           =   3075
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CDDADA&
      Caption         =   "SELECCIÓN"
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
      Left            =   3600
      TabIndex        =   23
      Top             =   2700
      Width           =   3315
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
         Left            =   3060
         TabIndex        =   32
         Top             =   1080
         Width           =   175
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
         Height          =   285
         Left            =   2610
         TabIndex        =   31
         Top             =   1080
         Width           =   435
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
         Height          =   285
         Left            =   2610
         TabIndex        =   25
         Top             =   270
         Width           =   435
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
         Height          =   285
         Left            =   3030
         TabIndex        =   24
         Top             =   270
         Width           =   175
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Todos - Sin Clasificación"
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
         Left            =   120
         TabIndex        =   34
         Top             =   1500
         Width           =   3135
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Deposito"
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
         Left            =   230
         TabIndex        =   33
         Top             =   1170
         Width           =   2325
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Material"
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
         Left            =   240
         TabIndex        =   28
         Top             =   360
         Width           =   2325
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Todos - Sin Clasificación"
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
         Left            =   90
         TabIndex        =   26
         Top             =   690
         Width           =   3135
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00CDDADA&
      Caption         =   "RANGO DE FECHAS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   60
      TabIndex        =   16
      Top             =   1785
      Width           =   4600
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
         Height          =   300
         Left            =   4215
         TabIndex        =   20
         Top             =   375
         Width           =   175
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
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
         Left            =   3225
         TabIndex        =   19
         Text            =   " "
         Top             =   375
         Width           =   1020
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
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
         Left            =   1095
         TabIndex        =   18
         Top             =   375
         Width           =   1020
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
         Left            =   2085
         TabIndex        =   17
         Top             =   375
         Width           =   175
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta"
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
         Left            =   2475
         TabIndex        =   22
         Top             =   480
         Width           =   645
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Desde"
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
         Left            =   -1320
         TabIndex        =   21
         Top             =   480
         Width           =   2325
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   7575
      Left            =   7080
      ScaleHeight     =   7545
      ScaleWidth      =   1665
      TabIndex        =   13
      Top             =   0
      Width           =   1695
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   105
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   42
         Top             =   5040
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   43
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command9 
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
         Picture         =   "ESTASTK07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Listado de Materiales Sin Movimientos"
         Top             =   1980
         Width           =   650
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Ex"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "ESTASTK07.frx":0614
         TabIndex        =   38
         ToolTipText     =   "Exportar Movimientos de Stock"
         Top             =   3675
         Width           =   650
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
         Height          =   690
         Left            =   105
         Picture         =   "ESTASTK07.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Stock Minimo Recomendado"
         Top             =   2810
         Width           =   650
      End
      Begin VB.CommandButton Command11 
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
         Left            =   105
         Picture         =   "ESTASTK07.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   865
         Width           =   650
      End
      Begin VB.CommandButton Command10 
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
         Height          =   690
         Left            =   105
         Picture         =   "ESTASTK07.frx":0F32
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   180
         Width           =   650
      End
   End
   Begin VB.ListBox List1 
      Height          =   540
      Left            =   9135
      Sorted          =   -1  'True
      TabIndex        =   8
      Top             =   2205
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00CDDADA&
      Caption         =   "RANGO DE CODIGOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1590
      Left            =   60
      TabIndex        =   1
      Top             =   105
      Width           =   6840
      Begin VB.CommandButton Command2 
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
         Left            =   6500
         TabIndex        =   2
         Top             =   460
         Width           =   175
      End
      Begin VB.CommandButton Command1 
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
         Left            =   3120
         TabIndex        =   4
         Top             =   460
         Width           =   175
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
         Height          =   300
         Left            =   760
         TabIndex        =   5
         Top             =   460
         Width           =   2350
      End
      Begin VB.TextBox Text2 
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
         Left            =   4120
         TabIndex        =   3
         Top             =   460
         Width           =   2350
      End
      Begin VB.Label Label2 
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
         Height          =   495
         Left            =   3465
         TabIndex        =   10
         Top             =   900
         Width           =   3255
      End
      Begin VB.Label Label1 
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
         Height          =   495
         Left            =   105
         TabIndex        =   9
         Top             =   900
         Width           =   3255
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Desde"
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
         Left            =   150
         TabIndex        =   7
         Top             =   560
         Width           =   525
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta"
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
         Left            =   3530
         TabIndex        =   6
         Top             =   560
         Width           =   525
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00CDDADA&
      Caption         =   "FORMATO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3600
      TabIndex        =   0
      Top             =   4620
      Width           =   2700
      Begin VB.OptionButton Option12 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Acumulado"
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
         Left            =   1440
         TabIndex        =   12
         Top             =   360
         Value           =   -1  'True
         Width           =   1140
      End
      Begin VB.OptionButton Option11 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Detalle"
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
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   930
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "ESTASTK07.frx":107C
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "ESTASTK07.frx":12B0
      TabIndex        =   40
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivos"
      Begin VB.Menu mnuInvertir 
         Caption         =   "Invertir Selección de Tipo de Movimiento"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu mnuCL 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu mnuEstadisMensual 
         Caption         =   "Estadístico de Stocks Por tipo de Movimiento Presentación Mensual"
      End
      Begin VB.Menu mnuEstadisticasSalidasMensual 
         Caption         =   "Estadisticas de Stocks Solo Salidas Presentación Mensual"
      End
      Begin VB.Menu mnuestaconsuencolummes 
         Caption         =   "Estadisticas de Stocks Solo Salidas Encolumnado Mensual"
      End
   End
End
Attribute VB_Name = "ESTASTK07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim COTIMOV$()



Private Sub Command1_Click()
    Call SELECHO("MASTER")
    COD1$ = TRIM$(VALACT1$("MASTER"))
    If COD1$ <> "" Then
       Text1.TEXT = COD1$
    End If
End Sub

Private Sub Command10_Click()
   Call CIERRARCH("*.*")
   Unload Me
End Sub

Private Sub Command12_Click()
    'DECLARACIONES
    Dim CAENT#(), CASAL#(), ACTI%()
    Command12.Enabled = False
    'VALIDACIONES
    If TRIM(Text4) = "" Or TRIM(Text5) = "" Then
      Call COMUNI("Rango de Fecha no valido")
      Command12.Enabled = True
      Exit Sub
    End If
    
    F11% = FECHANUM(TRIM(Text4))
    F22% = FECHANUM(TRIM(Text5))
    
    If F11% > F22% Then
      Call COMUNI("Rango de Fecha no valido")
      Command12.Enabled = True
      Exit Sub
    End If
    '
    If TRIM(Text1) = "" Or TRIM(Text2) = "" Then
      Call COMUNI("Código no Valido.")
      Command12.Enabled = True
      Exit Sub
    End If
    C1% = CODINT(Text1)
    C2% = CODINT(Text2)
    If C1% <= 0 Or C2% <= 0 Then
      Call COMUNI("Código no Valido.")
      Command12.Enabled = True
      Exit Sub
    End If
    '
    HAYCOMOV% = 0
    ReDim COTIMOV$(List2.ListCount - 1)
    For i% = 0 To List2.ListCount - 1
      COTIMOV$(i%) = ""
    Next i%
    
    For i% = 0 To List2.ListCount - 1
      If List2.Selected(i%) = True Then
        HAYCOMOV% = 1
        COTIMOV$(i%) = List2.List(i%)
        
      End If
    Next i%
    If HAYCOMOV% = 0 Then
      Call COMUNI("Código de Movimiento no Valido.")
      Command12.Enabled = True
      Exit Sub
    End If
    
    'ARMADO DE CONSULTA
    
    
    'Call ABRESTOCKS
    'Screen.MousePointer = 11
    'SQLX$ = "SELECT * FROM MOVISTO "
    'SQLX$ = SQLX$ + " WHERE CDBL( FechMov ) >= " & F1# & " "
    'SQLX$ = SQLX$ + " AND CDBL( FechMov ) <= " & F2# & " "
    
    ' SQL
    Call ABRESTOCKS
    Screen.MousePointer = 11
    
    
    SQLX$ = "SELECT ReferCor , COD_INTE, DEPOMOVI, FECHMOV, IdenLote, CodiMovi, DoPriCor, DoSecCor "
    SQLX$ = SQLX$ + ", VALOUNIT , FeReMovi, CANTSALID, CANTINGRE, Nume_Prov, Nume_Clie, MoneVal , NumUsuar "
    SQLX$ = SQLX$ + " FROM MOVISTO WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE FechMov >= '" & FETEXCOR1$(F11%) & "' "
    SQLX$ = SQLX$ + " AND FechMov <= '" & FETEXCOR1$(F22%) & "' "
    
'    SQLX$ = "SELECT ReferCor , COD_INTE, DEPOMOVI, FECHMOV, IdenLote, CodiMovi, DoPriCor, DoSecCor,MONTH(FechMov) as mes ,YEAR(C.FechMov) AS ANO "
'    SQLX$ = SQLX$ + ", VALOUNIT , FeReMovi, CANTSALID, CANTINGRE, Nume_Prov, Nume_Clie, MoneVal , NumUsuar "
'    SQLX$ = SQLX$ + " FROM MOVISTO WITH(NOLOCK) "
'    SQLX$ = SQLX$ + " WHERE FechMov >= '" & FETEXCOR1$(F11%) & "' "
'    SQLX$ = SQLX$ + " AND FechMov <= '" & FETEXCOR1$(F22%) & "' "
'    SQLX$ = SQLX$ + " group by  YEAR(C.FECHEMISI), MONTH(C.FECHEMISI),ReferCor , COD_INTE, DEPOMOVI, FECHMOV, IdenLote, CodiMovi, DoPriCor, DoSecCor"
'    SQLX$ = SQLX$ + " ordr by  MONTH(C.FECHEMISI), YEAR(C.FECHEMISI), "
    M1% = 0
    For i% = 0 To List2.ListCount - 1 '8\\\OT-06-0218-WAR-05/06/06///
      If COTIMOV$(i%) <> "" Then
        If M1% = 0 Then
          COTIMOV1$ = Left(COTIMOV$(i%), 2)
          SQLX$ = SQLX$ + " AND ( CodiMovi = '" & COTIMOV1$ & "'"
          M1% = 1
        Else
          COTIMOV1$ = Left(COTIMOV$(i%), 2)
          If COTIMOV1$ = "VM" Then
             If TRIM$(Control$("", "CODCONSU")) <> "" Then
                COTIMOV1$ = TRIM$(Control$("", "CODCONSU"))
             End If
          End If
          SQLX$ = SQLX$ + " OR CodiMovi = '" & COTIMOV1$ & "'  "
        End If
      End If
    Next i%
    If M1% > 0 Then SQLX$ = SQLX$ + " )"
    '
    If XVALO(Text6) > 0 Then 'filtro desposito
      SQLX$ = SQLX$ + "AND DepoMovi = " & XVALO(Text6) & " "
    End If
       '
    'RECORDSET
    'Set RST1TMP = Stocks.OpenRecordset(SQLX$, 4)
    ' SQL
'    Dim RST1TMP As ADODB.Recordset
'    Set RST1TMP = New ADODB.Recordset
'    RST1TMP.CursorType = adOpenDynamic
'    RST1TMP.LockType = adLockReadOnly
'    RST1TMP.Open FILTSQL$(SQLX$), FLEXCONN, , , adCmdText
    Set RST1TMP = READSET(SQLX$)
    If (RST1TMP.EOF And RST1TMP.BOF) Then ' si esta vacio
      Call COMUNI("No Existen Movimientos para el Criterio Seleccionado.")
      Screen.MousePointer = 11
      GoTo 99
    End If
    
    With RST1TMP
      'On Error Resume Next
      '.MoveLast
      'If Err Then
      '  Call COMUNI("No Existen Movimientos para el Criterio Seleccionado.")
      '  GoTo 99
      'End If
      
      ULT1& = .RecordCount
      .MoveFirst
      On Error GoTo 0
      
      
      'INICIALIZACIONES
      JJ& = 0
      CONT& = 0
      ReDim CAENT#(NUMAS%)
      ReDim CASAL#(NUMAS%)
      ReDim ACTI%(NUMAS%)
      For i% = 1 To NUMAS%
        CAENT#(i%) = 0
        CASAL#(i%) = 0
        ACTI%(i%) = 0
      Next i%
      Call COPYSTRU("COKARDEX", "LIMOVSTO")
      REBLA$ = REGBLAN$("COKARDEX")
      CODI1$ = TRIM(Text1)
      CODI2$ = TRIM(Text2)
      '
      SQLX$ = "SELECT Codint FROM MASTER WITH(NOLOCK) "
      SQLX$ = SQLX$ + " WHERE Codigo >= '" & CODI1$ & "' "
      SQLX$ = SQLX$ + " AND Codigo <= '" & CODI2$ & "' "
      If XVALO(Text3) > 0 Then
         SQLX$ = SQLX$ + " AND Timate = " & XVALO(Text3)
      End If
'      Set CODISEL = FLEXCONN.Execute(FILTSQL$(SQLX$))
      Set CODISEL = READSET(SQLX$)
      '
      With CODISEL
         If Not (.BOF And .EOF) Then
            Do While Not .EOF
               CIII% = !CODINT
               If CIII% > 0 Then
                  ACTI%(CIII%) = 1
               End If
            .MoveNext
            Loop
         End If
      End With
      '
      If Check1.Value = 1 Then         ' INCLUYE COMPONENTES MULTINIVEL
         FIN& = NUMAS%
         For CIXI% = 1 To NUMAS%
            UK& = CIXI%
            Call TRACE(20, UK&, FIN&)
            If ACTI%(CIXI%) = 1 Then
               Call EXPLOMUL(CIXI%)
               For KI% = 0 To CONTAMUL%
                  On Error Resume Next
                  ACTI%(CIK%(KI%)) = 2
                  On Error GoTo 0
               Next KI%
            End If
         Next CIXI%
         For CIXI% = 1 To NUMAS%
            If ACTI%(CIXI%) > 0 Then ACTI%(CIXI%) = 1
         Next CIXI%
      End If
      '
'      FIN& = NUMAS%
'      For CIXI% = 1 To NUMAS%
'        UK& = CIXI%
'        Call TRACE(20, UK&, FIN&)
'        CODD$ = CODEXT$(CIXI%)
'        If CODD$ >= CODI1$ Then
'          If CODD$ <= CODI2$ Then
'            If XVALO(Text3) = 0 Or TIMATE%(CIXI%) = XVALO(Text3) Then 'filtro tipo material
'              '
'              ACTI%(CIXI%) = 1
'              '
'              If Check1.Value = 1 Then
'                Call EXPLOMUL(CIXI%)
'                For KI% = 0 To CONTAMUL%
'                  On Error Resume Next
'                  ACTI%(CIK%(KI%)) = 1
'                  On Error GoTo 0
'                Next KI%
'              End If
'              '
'            End If
'          End If
'        End If
'      Next CIXI%
      '
      'BUCLE
      Do While Not .EOF
        CIXI% = !cod_inte
        If CIXI% < 1 Or CIXI% > NUMAS% Then GoTo 15
        '
        If ACTI%(CIXI%) = 1 Then
          CODD$ = CODEXT$(CIXI%)
          '
          If XVALO(Text6) > 0 Then 'filtro desposito
            If !depomovi <> XVALO(Text6) Then
              GoTo 15
             End If
          End If
          '
          YY$ = REBLA$
          Call REPLA(YY$, MKI$(CVINT(!FechMov)), 1, 2)
          Call REPLA(YY$, MKI$(!cod_inte), 3, 2)
          Call REPLA(YY$, SAFETEXT$(!idenlote), 5, 16)
          Call REPLA(YY$, !CODIMOVI, 21, 2)
          Call REPLA(YY$, SAFETEXT$(!dopricor), 23, 10)
          '/** Por si no existe documento primario, pongo el secundario
          If TRIM$(!dopricor) = "" Then
            Call REPLA(YY$, SAFETEXT$(!DoSecCor), 23, 10)
          End If
          '**/
          Call REPLA(YY$, SAFETEXT$(!DoSecCor), 33, 10)
          Call REPLA(YY$, SAFETEXT$(!ReferCor), 43, 30)
          Call REPLA(YY$, MKD$(!VALOUNIT), 73, 8)
            MONEX$ = "$": If !MoneVal = 2 Then MONEX$ = "U$"
          Call REPLA(YY$, MONEX$, 81, 2)
          Call REPLA(YY$, Chr$(!depomovi), 83, 1)
          Call REPLA(YY$, MKI$(!NUME_CLIE), 84, 2)
          Call REPLA(YY$, MKI$(!Nume_Prov), 86, 2)
          Call REPLA(YY$, MKD$(!CANTINGRE), 88, 8)
            CAENT#(CIXI%) = CAENT#(CIXI%) + !CANTINGRE
          Call REPLA(YY$, MKD$(!CANTSALID), 96, 8)
            CASAL#(CIXI%) = CASAL#(CIXI%) + !CANTSALID
          Call REPLA(YY$, MKD$(0), 104, 8)
          Call REPLA(YY$, CVINT(!FeReMovi), 122, 2)
          Call REPLA(YY$, Chr$(XVALO(!NumUsuar)), 126, 1)
            JJ& = JJ& + 1
          'Call GRAREG("LIMOVSTO", YY$, JJ&)
          Call GRAREG("#LIMOVSTO", YY$, JJ&)
        End If
        '
15      CONT& = CONT& + 1
        Call TRACE(20, CONT&, ULT1&)
       .MoveNext
      Loop
    End With
    RST1TMP.Close
    Set RST1TMP = Nothing
    If JJ& = 0 Then
      Call COMUNI("No Existen Movimientos para el Criterio Seleccionado.")
      Command12.Enabled = True
      Screen.MousePointer = 11
      Exit Sub
    End If
    If Option12 Then 'SI ACUMULADO
      ZZ$ = REGBLAN("LIMOVSTO")
      JJ& = 0: FIN& = NUMAS%
      For i% = 1 To NUMAS%
        U& = i%
        Call TRACE(20, U&, FIN&)
        If CAENT#(i%) > 0 Or CASAL#(i%) > 0 Then
             Call REPLA(ZZ$, MKI$(i%), 3, 2)
             Call REPLA(ZZ$, MKD$(CAENT(i%)), 88, 8)
             Call REPLA(ZZ$, MKD$(CASAL(i%)), 96, 8)
             JJ& = JJ& + 1
             'Call GRAREG("LIMOVSTO", ZZ$, JJ&)
             Call GRAREG("#LIMOVSTO", ZZ$, JJ&)
        End If
      Next i%
    End If
    'CIERRE, ORDEN Y LLAMDO A ARCHIVO.
    'Call SETULTREG("LIMOVSTO", JJ&)
    Call SETULTREG("#LIMOVSTO", JJ&)
    Call FILEDUMP("LIMOVSTO")
    '
    Call FILESORT("LIMOVSTO", "", 1, 1, "ASC")     ' orden por fecha
    Call FILESORT("LIMOVSTO", "", 2, 2, "ASC")     ' orden por codigo
    If Option12 Then 'SI ACUMULADO
      Call FINAL("*LIMOSTAC")
    Else
      Call FINAL("*LIMOVSTO")
    End If
    Screen.MousePointer = 1
    '
99  On Error GoTo 0
    Command12.Enabled = True
    Screen.MousePointer = 1
End Sub

Private Sub Command2_Click()
    Call SELECHO("MASTER")
    COD2$ = TRIM$(VALACT1$("MASTER"))
    If COD2$ <> "" Then
       Text2.TEXT = COD2$
    End If
End Sub

Private Sub Command3_Click()
    Call SELECHO("TATIMAT")
    VDEV$ = VALACT1$("TATIMAT")
    If VDEV$ <> "" Then
       Text3.TEXT = VDEV$
       If TRIM$(Text3.TEXT) = "" Then
            Label4.Caption = "Todos - Sin Clasificación"
          Else
            Label4.Caption = VALACT2$("TATIMAT")
       End If
    End If
End Sub

Private Sub Command4_Click()
    Call SELECHO("SELFECHA")
    VDEV$ = VALACT1$("SELFECHA")
    If VDEV$ <> "" Then
        Text4.TEXT = VDEV$
    End If
End Sub

Private Sub Command43_Click()

End Sub

Private Sub Command5_Click()
    Call SELECHO("SELFECHA")
    VDEV$ = VALACT1$("SELFECHA")
    If VDEV$ <> "" Then
        Text5.TEXT = VDEV$
    End If
End Sub

Private Sub Command6_Click()
    Call SELECHO("TADEPOSI")
    VDEV$ = VALACT1$("TADEPOSI")
    If VDEV$ <> "" Then
       Text6.TEXT = VDEV$
       If TRIM$(Text6.TEXT) = "" Then
            Label8.Caption = "Todos - Sin Clasificación"
          Else
            Label8.Caption = VALACT2$("TADEPOSI")
       End If
    End If
End Sub

Private Sub Command7_Click()
  '\\\OT-05-0480-WAR-08/07/08///
  Dim CANTISAL#()
  
  Command7.Enabled = False
  Screen.MousePointer = 11
  Call DESHASFECHA(Des%, Has%, PERIO$)
  If PERIO$ = "" Then
    Screen.MousePointer = 1:
    Command7.Enabled = True
    Exit Sub
  End If
  Desde = Des%: Hasta = Has%
  CANDIA% = DIENTRE%(Desde, Hasta)
  'DES1& = CDbl(CVDAT(DES%))
  'HAS1& = CDbl(CVDAT(HAS%))
  DES1 = FETEXCOR1$(Des%)
  HAS1 = FETEXCOR1$(Has%)
  
9 Y2& = XVALO(Control("", "CANDIACO"))
  Y3$ = InputBox("Ingrese la Cantidad de Dias de Cobertura", "Stock Mínimo Recomendado", Y2&)
  Y4 = XVALO(Y3$)
  If Y4 > 0 And Y4 < 30000 Then
      Call GRACONTROL("", "CANDIACO", Y3$)
      CANDIACOB% = Y4 'CANTIDAD DE DIAS DE COBERTURA
    Else
      GoTo 99   ' Call COMUNI("Cantidad de Dias no Valida")
  End If
  '
  ReDim CANTISAL#(NUMAS%)
  For i% = 1 To NUMAS%
    CANTISAL#(i%) = 0
  Next i%
  '
  Call ABRESTOCKS
  SQLX$ = "SELECT COD_INTE, CANTSALID FROM MOVISTO  WITH(NOLOCK) "
  SQLX$ = SQLX$ + " WHERE CodiMovi <> 'AJ' "
  SQLX$ = SQLX$ + " AND CantSalid > 0 "
  SQLX$ = SQLX$ + "AND FechMov >= '" & DES1 & "' "
  SQLX$ = SQLX$ + "AND FechMov <= '" & HAS1 & "' "
  'Set SMRTMP = Stocks.OpenRecordset(sqlx$, 4)
  '
'  Dim SMRTMP As ADODB.Recordset
'  Set SMRTMP = New ADODB.Recordset
'  SMRTMP.CursorType = adOpenDynamic
'  SMRTMP.LockType = adLockReadOnly
'  SMRTMP.Open FILTSQL$(SQLX$), FLEXCONN, , , adCmdText
  Set SMRTMP = READSET(SQLX$)
  If (SMRTMP.EOF And SMRTMP.BOF) Then 'si esta vacio
      Call MENSERR(24, "Ningún Registro Seleccionado según Condición SQL")
      GoTo 99
  End If
  With SMRTMP
  '  On Error Resume Next
  '  .MoveFirst
  '  If Err Then
  '    Call MENSERR(24, "Ningún Registro Seleccionado según Condición SQL")
  '    GoTo 99
  '  End If
    On Error GoTo 0
    '
    Do While .EOF = False
      CIXI% = XVALO(!cod_inte)
      On Error Resume Next    ' POR SI CIXI% > NUMAS%
      CANTISAL#(CIXI%) = CANTISAL#(CIXI%) + XVALO(!CANTSALID)
      On Error GoTo 0
    .MoveNext
    Loop
  End With
  SMRTMP.Close
  Set SMRTMP = Nothing

  '
  For i% = 1 To NUMAS% 'DIVIDO POR LA CANTIDAD DE DIAS
    CANTISAL#(i%) = (CANTISAL#(i%) / CANDIA%) * CANDIACOB%
  Next i%
  '
  CI1% = PRICOD%
  CI2% = ULTCOD%
  VDEF$ = MKI$(CI1%) + MKI$(CI2%) + Chr$(0)
  '
10 OBX$ = OBJPLA$("DESHASCOD1", VDEF$)
   If Len(OBX$) < 5 Then
       GoTo 99
   End If
   '
   CO1% = CVI(Mid$(OBX$, 1, 2))
   CO2% = CVI(Mid$(OBX$, 3, 2))
   TMT% = Asc(Mid$(OBX$, 5, 1))
   COD1$ = CODEXT$(CO1%)
   COD2$ = CODEXT$(CO2%)
   TIMAT$ = ECOARCH$("TATIMAT", Chr$(TMT%))
   If TIMAT$ = "" Then
     TMT% = 0
     TIMAT$ = "Sin Clasificación"
   End If
   '
   Call HEADERS("STMINREC", "")
   Call HEADERS("STMINREC", "Tipo Material: " & TIMAT$)
   REBLA$ = REGBLAN("STMINREC")
   J& = 0
   T1& = XVALO(NUMAS%)
   For CIXI% = 1 To NUMAS%
     CDXX$ = CODEXT$(CIXI%)
     If CDXX$ >= COD1$ Then
       If CDXX$ <= COD2$ Then
         If TMT% = 0 Or TIMATE%(CIXI%) = TMT% Then 'filtro tipo material
            '
            ZZ$ = REBLA$
            Call REPLA(ZZ$, MKI$(CIXI%), 1, 2)
            Call REPLA(ZZ$, Unimed(CIXI%), 3, 2)
            Call REPLA(ZZ$, STOMIN(CIXI%), 5, 4)
            Call REPLA(ZZ$, MKS(CANTISAL#(CIXI%)), 9, 4)
            Call REPLA(ZZ$, MKS(LOTESTAN(CIXI%)), 13, 4)
            If STOMIN(CIXI%) < CANTISAL#(CIXI%) Then
              Call REPLA(ZZ$, "Ajustar St.Minimo", 21, 28)
            End If
            J& = J& + 1
            Call GRAREG("STMINREC", ZZ$, J&)
            '
         End If
       End If
     End If
     J1& = XVALO(CIXI%)
   Call TRACE(24, J1&, T1&)
11 Next CIXI%
   Call SETULTREG("STMINREC", J&)
   Call FINAL("*STMINRE1")
   GoTo 10
   '
99 On Error GoTo 0
   Screen.MousePointer = 1
   Command7.Enabled = True
   '\\\OT-05-0480-WAR-FIN///
   
   
End Sub

Private Sub Command8_Click()
Call MENSERR(24, "No Disponible en Versión SQL-Server")
Exit Sub
   '
   Command8.Enabled = False
   Call DESHASFECHA(Des%, Has%, PERIO$)
   If PERIO$ <> "" Then
10   If EXISTE%(LUCOM$ + "STOCKS.MDB") > 0 Then
        Kill LUCOM$ + "STOCKS.MDB"
        GoTo 10
     End If
     '
     Call APERBASDAT("STOCKS")
     Set CREABASTOK = wrkPredeterminado.CreateDatabase(LUCOM$ + "STOCKS.MDB", dbLangGeneral)
     Dim BASTOCK As Database
     Set BASTOCK = OpenDatabase(LUCOM$ + "STOCKS.MDB")
     '
     Dim MOVIST As New TableDef
     
     Set CREATABSTO = BASTOCK.CreateTableDef("MOVISTO")
     NAMX = Movisto.Fields(0).Name
     TIPX = Movisto.Fields(0).Type
     SIZX = Movisto.Fields(0).Size
     CREATABSTO.Fields.Append CREATABSTO.CreateField(NAMX, TIPX, SIZX) ', TIPOVA, LOCAM%)
     BASTOCK.TableDefs.Append CREATABSTO
     Set MOVIST = BASTOCK.TableDefs("MOVISTO")
     '
     FIN& = Movisto.Fields.Count
     For KK1% = 2 To Movisto.Fields.Count
       NAMX = Movisto.Fields(KK1% - 1).Name
       TIPX = Movisto.Fields(KK1% - 1).Type
       SIZX = Movisto.Fields(KK1% - 1).Size
       With MOVIST.Fields
         .Append MOVIST.CreateField(NAMX, TIPX, SIZX)
         TICAM% = TIPX
         NCAMPO$ = NAMX
         If TICAM% = 1 Or TICAM% = 10 Or TICAM% = 12 Then
               MOVIST.Fields(NCAMPO$).AllowZeroLength = True
            ElseIf TICAM% = 1 Then
               MOVIST.Fields(NCAMPO$).DefaultValue = True
            ElseIf TICAM% = 8 Then    ' FECHA
               MOVIST.Fields(NCAMPO$).DefaultValue = CVDAT(0)
            Else
               MOVIST.Fields(NCAMPO$).DefaultValue = 0
         End If
       End With
     Next KK1%
     
     Set MOVISLOC = BASTOCK.OpenRecordset("MOVISTO", dbOpenDynaset)
     '
     F11$ = FETEXCOR1$(Des%)
     F22$ = FETEXCOR1$(Has%)
     '
     'ARMADO DE CONSULTA
     Call ABRESTOCKS
     Screen.MousePointer = 11
     SQLX$ = "SELECT * FROM MOVISTO "
     SQLX$ = SQLX$ + " WHERE FechMov  >= '" & F11$ & "' "
     SQLX$ = SQLX$ + " AND  FechMov <= '" & F22$ & "' "
     Set RST1TMP = Stocks.OpenRecordset(SQLX$, 4)
     With RST1TMP
       On Error Resume Next
       .MoveLast
       If Err Then
         Call COMUNI("No Existen Movimientos para el Criterio Seleccionado.")
         GoTo 99
       End If
       ULT1& = .RecordCount
       .MoveFirst
       KKII& = 0
       On Error GoTo 0
       '
       Do While Not .EOF
          MOVISLOC.AddNew   ' YA CONVERTIDO
                  On Error Resume Next
          MOVISLOC!CODIEMPR = CodiEmp%
                  On Error GoTo 0
          For KKI% = 1 To .Fields.Count
            XVA = .Fields(KKI% - 1).Value
            MOVISLOC.Fields(KKI% - 1).Value = XVA
          Next KKI%
          MOVISLOC.Update
       .MoveNext
       KKII& = KKII& + 1
       Call TRACE(20, KKII&, ULT1&)
       Loop
     End With
     '
   End If
   '
   Call COMUNI("Base de Datos 'STOCKS.MDB'\Creada en 'C:/FLEXOFT'\Con " + TRIM$(Str$(ULT1&)) + " Registros.")
   '
99 Screen.MousePointer = 1
   Command8.Enabled = True
End Sub


Private Sub Command9_Click()
   Command9.Enabled = False
   
5  OPX% = COMALTER%("Articulos Sin Movimiento\\Por Rango de Fechas\Por Rango de Fechas y Con Estructura")
   If OPX% < 1 Or OPX% > 2 Then GoTo 99
   
   If OPX% = 1 Then
        Call MATSINMOV ' listado de articulos sin movimiento
        GoTo 5
   Else
        Call MATSINMOV_CONEXPLOSION
        GoTo 5
   End If
   
99 Command9.Enabled = True
End Sub

Private Sub Form_Load()

UTILIZA_SKIN% = 1
Call APLICACIONSKINS(Me)

'
List2.AddItem "AJ - Ajustes"
List2.AddItem "BR - Recepción"
List2.AddItem "CF - Consumo Fábrica"
List2.AddItem "DV - Devolución"
List2.AddItem "OE - Orden de Empaque"
List2.AddItem "OF - Orden de Fabricación"
List2.AddItem "PF - Parte de Fabricación"
List2.AddItem "RC - Remito de Consignación"
List2.AddItem "RT - Remito de Ventas"
List2.AddItem "TR - Transferencias de Stock"
List2.AddItem "SC - Vale de Scrap"
List2.AddItem "VM - Vale de Consumo"
List2.AddItem "RD - Remito Devolución"
List2.AddItem "RP - Consumo Reparación"
'
'Valores por default
Text4 = FECHATEX$(COMESACT%)
Text5 = HOS$
'
'\\\OT-08-0990-OD-30/10/08///
LOMA% = XVALO(Control$("MASTER", "LOMAXCOD"))
If LOMA% < 15 Or LOMA% > 24 Then LOMA% = 15
If LOMA% > 15 Then
    Text1.MaxLength = LOMA%
    Text2.MaxLength = LOMA%
End If
'
Text1 = TRIM$(CODEXT$(PRICOD%))
Text2 = TRIM(CODEXT$(ULTCOD%))
'
End Sub

Private Sub mnuestaconsuencolummes_Click()
    If TRIM(Text4) = "" Or TRIM(Text5) = "" Then
      Call COMUNI("Rango de Fecha no valido")
     
      Exit Sub
    End If
    
    Des% = FECHANUM(TRIM(Text4))
    Has% = FECHANUM(TRIM(Text5))
    
    If Des% > Has% Then
      Call COMUNI("Rango de Fecha no valido")
      Exit Sub
    End If
    '
    If TRIM(Text1) = "" Or TRIM(Text2) = "" Then
      Call COMUNI("Código no Valido.")
      Command12.Enabled = True
      Exit Sub
    End If
    C1% = CODINT(Text1)
    C2% = CODINT(Text2)
    If C1% <= 0 Or C2% <= 0 Then
      Call COMUNI("Código no Valido.")
      
      Exit Sub
    End If

    If IsFormLoaded(FRMZELECHO) Then
        ESTAENMEMORIA = 1
    Else
        Call HookForm(FRMZELECHO.msf2) 'estaba tirando un error cuando se cierra el formulario con doble clic
        Call HookForm(FRMZELECHO.MSF1) 'probablemente por que hace un hide, SI NO ESTA EN MEMORIA SE APLICA LA FUNCION.
    End If
    '
    ReDim COTIMOV$(List2.ListCount - 1)
    For i% = 0 To List2.ListCount - 1
      COTIMOV$(i%) = ""
    Next i%
    
    ReDim COTIMOV$(List2.ListCount - 1)
    For i% = 0 To List2.ListCount - 1
      COTIMOV$(i%) = ""
    Next i%
    
    For i% = 0 To List2.ListCount - 1
      If List2.Selected(i%) = True Then
        HAYCOMOV% = 1
        COTIMOV$(i%) = List2.List(i%)
      End If
    Next i%
    
    If HAYCOMOV% = 0 Then
      Call COMUNI("Código de Movimiento no Valido.")
      Exit Sub
    End If
    M1% = 0
    For i% = 0 To List2.ListCount - 1 '8\\\OT-06-0218-WAR-05/06/06///
      If COTIMOV$(i%) <> "" Then
        If M1% = 0 Then
          COTIMOV1$ = Left(COTIMOV$(i%), 2)
          SQLX$ = SQLX$ + " AND ( CodiMovi = '" & COTIMOV1$ & "'"
          M1% = 1
        Else
          COTIMOV1$ = Left(COTIMOV$(i%), 2)
          If COTIMOV1$ = "VM" Then
             If TRIM$(Control$("", "CODCONSU")) <> "" Then
                COTIMOV1$ = TRIM$(Control$("", "CODCONSU"))
             End If
          End If
          SQLX$ = SQLX$ + " OR CodiMovi = '" & COTIMOV1$ & "'  "
        End If
      End If
    Next i%
    If M1% > 0 Then SQLX$ = SQLX$ + " )"
    '
    SQLCODMOV$ = SQLX$
    If TRIM$(SQLCODMOV$) = "" Then
      Call COMUNI("Falta Seleccionar Tipo de Movimiento")
    End If
    '
    Campos$ = "Código/a16;Descripción/a36"
    
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    '
    'CARGA LOS ENCABEZADOS CON LOS MESES DENTRO DEL RANGO A CONTINUACION DE LOS CARGADOS CON CARGA_ENCABEZADO.
    Call CARGA_ENCABEZADO_MESES(FRMZELECHO.msf2, FRMZELECHO, Des%, Has%, "F12.1", 3, RET_QMESES%) 'RET_QMESES RETORNA LA CANT.MESES
    Call CARGA_ENCABEZADO_MESES(FRMZELECHO.MSF1, FRMZELECHO, Des%, Has%, "F12.1", 3, RET_QMESES%) 'DENTRO DEL RANGO
    '
    FRMZELECHO.Caption = "ESTADÍSTICO DE CONSUMOS FORMATO MENSUAL SOLO MOV. DE SALIDA"
    DEP% = 0: If XVALO(Text6) > 0 Then DEP% = XVALO(Text6)
    TMAT% = 0: If XVALO(Text3) > 0 Then TMAT% = XVALO(Text3)
    Dim obj As New RECORXET
    Set RS1 = obj.RS_ESTADISTICA_CONSUMO_AGRUPADA_X_MES_ENCOLUMNADO(Des%, Has%, DEP%, SQLCODMOV$, TMAT%, C1%, C2%, 1)
    VUELTA% = 0: J& = 0: VUELTITA% = 0
    CantidadMesesConConsumo% = 0
    With RS1
      Screen.MousePointer = 11
      Do While Not .EOF
         MESLEIDO$ = UCase$(Left$(SAFETEXT$(!MES), 3))
         MESNUMLEIDO% = XVALO(!MES1)
         ANOLEIDO% = XVALO(!ANO)
         CODILEIDO$ = SAFETEXT$(!Cod_Exte)
         MESYANO$ = MESLEIDO$ & ". " & SAFETEXT$(ANOLEIDO%)
         COLUM% = NRODECOLUMNA_SEGUN_TEXTO%(FRMZELECHO.msf2, MESYANO$, 9) 'PARA MES 9 'SEP. 2019'
         'solicito que se agregue el Stock
         If VUELTITA% = 0 Then
           With FRMZELECHO.msf2
                .Cols = .Cols + 1
                .TextMatrix(0, .Cols - 1) = "Prom.(meses c-cons)": .ColWidth(.Cols - 1) = 2200: .ColAlignment(.Cols - 1) = flexAlignRightTop
                .Cols = .Cols + 1
                .TextMatrix(0, .Cols - 1) = "Promedio": .ColWidth(.Cols - 1) = 1100: .ColAlignment(.Cols - 1) = flexAlignRightTop
                .Cols = .Cols + 1
                .TextMatrix(0, .Cols - 1) = "Stock": .ColWidth(.Cols - 1) = 1100: .ColAlignment(.Cols - 1) = flexAlignRightTop
                .Cols = .Cols + 1
                .TextMatrix(0, .Cols - 1) = "C.Calculado": .ColWidth(.Cols - 1) = 1300: .ColAlignment(.Cols - 1) = flexAlignRightTop
                .Cols = .Cols + 1
                .TextMatrix(0, .Cols - 1) = "Stock Min": .ColWidth(.Cols - 1) = 1300: .ColAlignment(.Cols - 1) = flexAlignRightTop
                .Cols = .Cols + 1
                .TextMatrix(0, .Cols - 1) = "Lead Time": .ColWidth(.Cols - 1) = 1300: .ColAlignment(.Cols - 1) = flexAlignRightTop
                VUELTITA% = 1
           End With
           With FRMZELECHO.MSF1
                .Cols = .Cols + 1
                .TextMatrix(0, .Cols - 1) = "Prom.(meses c-cons)": .ColWidth(.Cols - 1) = 2200: .ColAlignment(.Cols - 1) = flexAlignRightTop
                .Cols = .Cols + 1
                .TextMatrix(0, .Cols - 1) = "Promedio": .ColWidth(.Cols - 1) = 1100: .ColAlignment(.Cols - 1) = flexAlignRightTop
                .Cols = .Cols + 1
                .TextMatrix(0, .Cols - 1) = "Stock": .ColWidth(.Cols - 1) = 1100: .ColAlignment(.Cols - 1) = flexAlignRightTop
                .Cols = .Cols + 1
                .TextMatrix(0, .Cols - 1) = "C.Calculado": .ColWidth(.Cols - 1) = 1300: .ColAlignment(.Cols - 1) = flexAlignRightTop
                .Cols = .Cols + 1
                .TextMatrix(0, .Cols - 1) = "Stock Min": .ColWidth(.Cols - 1) = 1300: .ColAlignment(.Cols - 1) = flexAlignRightTop
                .Cols = .Cols + 1
                .TextMatrix(0, .Cols - 1) = "Lead Time": .ColWidth(.Cols - 1) = 1300: .ColAlignment(.Cols - 1) = flexAlignRightTop
           End With
         End If
         'SI CAMBIA EL CODIGO ESCRIBE CODIGO Y DESCRIPCION
         If CODILEIDO$ <> COD_ANT$ Then
            If J& > 0 Then
                If RET_QMESES% > 0 Then
                    sPromedioConsumo$ = FORMATNUM$(ConsumoTotal# / RET_QMESES%, "F10.1")
                End If
                FRMZELECHO.msf2.TextMatrix(J&, FRMZELECHO.msf2.Cols - 5) = sPromedioConsumo$
                If CantidadMesesConConsumo% > 0 Then
                    sPromedioConConsumo$ = FORMATNUM$(ConsumoTotal# / CantidadMesesConConsumo%, "F10.1")
                End If
                FRMZELECHO.msf2.TextMatrix(J&, FRMZELECHO.msf2.Cols - 6) = sPromedioConConsumo$
            End If
            '
            J& = J& + 1
            FRMZELECHO.msf2.Rows = J& + 1
            FRMZELECHO.msf2.TextMatrix(J&, 1) = SAFETEXT$(!Cod_Exte)
            FRMZELECHO.msf2.TextMatrix(J&, 2) = SAFETEXT$(!Descrip)
            CI1% = XVALO(!cod_inte)
            ST$ = FORMATNUM$(STOCKACT(CI1%), "F10.1")
            FRMZELECHO.msf2.TextMatrix(J&, FRMZELECHO.msf2.Cols - 4) = ST$
            CostoCalculado$ = FORMATNUM$(!COSCAL, "F12.4")
            FRMZELECHO.msf2.TextMatrix(J&, FRMZELECHO.msf2.Cols - 3) = CostoCalculado$
            StockMinimo$ = FORMATNUM$(!STOMIN, "F10.1")
            FRMZELECHO.msf2.TextMatrix(J&, FRMZELECHO.msf2.Cols - 2) = StockMinimo$
            LEADTIME$ = FORMATNUM$(!LETICOM, "F10.1")
            FRMZELECHO.msf2.TextMatrix(J&, FRMZELECHO.msf2.Cols - 1) = LEADTIME$
                        
            ConsumoPeriodo# = XVALO(!SALE)
            CantidadMesesConConsumo% = 0
            If ConsumoPeriodo# > 0.05 Then CantidadMesesConConsumo% = CantidadMesesConConsumo% + 1
            ConsumoTotal# = ConsumoPeriodo#
         Else
            ConsumoPeriodo# = XVALO(!SALE)
            If ConsumoPeriodo# > 0.05 Then CantidadMesesConConsumo% = CantidadMesesConConsumo% + 1
            ConsumoTotal# = ConsumoTotal# + ConsumoPeriodo#
         End If
         'EN LA MISMA FILA SEGUN EL MES Y AÑO VA ESCRIBIENDO EN LA MISMA FILA SEGUN COLUMNA CORRESPONDIENTE.
         FRMZELECHO.msf2.TextMatrix(J&, COLUM%) = FORMATNUM$(XVALO(!SALE), "F12.1")
            
         MES_ANT$ = SAFETEXT$(!MES)
         MESNUM_ANT% = XVALO(!MES1)
         COD_ANT$ = CODILEIDO$
        .MoveNext
      Loop
    End With
    If J& > 0 Then
        If RET_QMESES% > 0 Then
            sPromedioConsumo$ = FORMATNUM$(ConsumoTotal# / RET_QMESES%, "F10.1")
        End If
        FRMZELECHO.msf2.TextMatrix(J&, FRMZELECHO.msf2.Cols - 3) = sPromedioConsumo$
        If CantidadMesesConConsumo% > 0 Then
            sPromedioConConsumo$ = FORMATNUM$(ConsumoTotal# / CantidadMesesConConsumo%, "F10.1")
        End If
        FRMZELECHO.msf2.TextMatrix(J&, FRMZELECHO.msf2.Cols - 4) = sPromedioConConsumo$
    End If
    
    
'    Call Carga_MSF_Recordset(RS1, CAMPOS$, FRMZELECHO.msf2, 1)
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    Screen.MousePointer = 1
    FRMZELECHO.BorderStyle = 2
    FRMZELECHO.mnuImprimir.Visible = True
'    FRMZELECHO.mnuExpExcel.Visible = True
    
    FRMZELECHO.Width = 13500
    FRMZELECHO.msf2.FixedCols = 3 ' PARA QUE EL CODIGO Y LA DESCRIPCION
    FRMZELECHO.MSF1.FixedCols = 3 ' QUEDE FIJO AL SCROLEAR
    FRMZELECHO.Font.Size = 8
    FRMZELECHO.Show 1
    Dim MSFX As msFlexGrid
    Set MSFX = FRMZELECHO.msf2
    If FRMZELECHO.MSF1.Rows > 1 Then Set MSFX = FRMZELECHO.MSF1
    
     If FRMZELECHO.EXPORTAREXCEL% > 0 Then
        CommonDialog1.CancelError = True
        CommonDialog1.DialogTitle = "Exportar a Excel"
            ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
        CommonDialog1.Filter = "Excel Files(*.XLS)|*.XLS|"
        CommonDialog1.FilterIndex = 1
        On Error Resume Next
        CommonDialog1.ShowOpen
        On Error GoTo 0
        If Err Then
           Exit Sub
        End If
        '
        DoEvents
        Screen.MousePointer = 11
        AREXPO$ = UCase$(CommonDialog1.FileName)
        If TRIM$(AREXPO$) <> "" Then
           'esta tirando un error en esta rutina
           Call GeneraExcelDeGrilla(AREXPO$, MSFX, "Estadística de Consumo")
           Exit Sub
        End If
     End If

    
    If FRMZELECHO.IMPRIMIR > 0 Then
'     If RET_QMESES% > 16 Then
'       Call COMUNI("Imposible Imprimir Reporte Cantidad de Campos Excede el Limite")
'       Exit Sub
'     End If
     If FRMZELECHO.MSF1.Rows > 1 Then Set MSFX = FRMZELECHO.MSF1
     If MSFX.Cols > 30 Then
        Call COMUNI("Imposible Imprimir Reporte Cantidad de Columnas Excede el Limite\Recomendamos Exportar a Excel")
        Exit Sub
     End If
     
     
'     CantidadDeRegistros& = XVALO(MSFX.Rows)
'     TOTALCELDAS& = CantidadDeRegistros& * (RET_QMESES% + 8)
'     If TOTALCELDAS& > 30000 Then
'        Call COMUNI("Imposible Imprimir Reporte Cantidad de Celdas Excede el Limite\Recomendamos Exportar a Excel")
'        Exit Sub
'     End If
     
     Campos$ = Campos$ & ";" & STRINGFILA$(MSFX, 0, 3, "/F12.1") ' LE CONCATENO LOS CAMPOS DE MESES QUE SE AGREGARON AUTOMATICAMENTE.
     If Right(Campos$, 1) = ";" Then Campos$ = Mid$(Campos$, 1, Len(Campos$) - 1)
     'CAMPOS$ = CAMPOS$ & ";C.Calculado/F12.4"
       
     Call CARGA_TABLA_IMPRE(MSFX, Campos$, "ESTADMES", "ESTADÍSTICO DE CONSUMOS FORMATO MENSUAL", "TABLA_IMPRE", RET_NAMECONS$, 1)
     Call FINAL("?" & RET_NAMECONS$)
'     Call EliminarSQF(RET_NAMECONS$)
    End If

End Sub




Private Sub mnuEstadisMensual_Click()
    If TRIM(Text4) = "" Or TRIM(Text5) = "" Then
      Call COMUNI("Rango de Fecha no valido")
     
      Exit Sub
    End If
    
    Des% = FECHANUM(TRIM(Text4))
    Has% = FECHANUM(TRIM(Text5))
    
    If Des% > Has% Then
      Call COMUNI("Rango de Fecha no valido")
      Exit Sub
    End If
    '
    If TRIM(Text1) = "" Or TRIM(Text2) = "" Then
      Call COMUNI("Código no Valido.")
      Command12.Enabled = True
      Exit Sub
    End If
    C1% = CODINT(Text1)
    C2% = CODINT(Text2)
    If C1% <= 0 Or C2% <= 0 Then
      Call COMUNI("Código no Valido.")
      
      Exit Sub
    End If

    If IsFormLoaded(FRMZELECHO) Then
        ESTAENMEMORIA = 1
    Else
        Call HookForm(FRMZELECHO.msf2) 'estaba tirando un error cuando se cierra el formulario con doble clic
        Call HookForm(FRMZELECHO.MSF1) 'probablemente por que hace un hide, SI NO ESTA EN MEMORIA SE APLICA LA FUNCION.
    End If
    '
    ReDim COTIMOV$(List2.ListCount - 1)
    For i% = 0 To List2.ListCount - 1
      COTIMOV$(i%) = ""
    Next i%
    
    For i% = 0 To List2.ListCount - 1
      If List2.Selected(i%) = True Then
        HAYCOMOV% = 1
        COTIMOV$(i%) = List2.List(i%)
      End If
    Next i%
    
    If HAYCOMOV% = 0 Then
      Call COMUNI("Código de Movimiento no Valido.")
      Command12.Enabled = True
      Exit Sub
    End If
    M1% = 0
    For i% = 0 To List2.ListCount - 1 '8\\\OT-06-0218-WAR-05/06/06///
      If COTIMOV$(i%) <> "" Then
        If M1% = 0 Then
          COTIMOV1$ = Left(COTIMOV$(i%), 2)
          SQLX$ = SQLX$ + " AND ( CodiMovi = '" & COTIMOV1$ & "'"
          M1% = 1
        Else
          COTIMOV1$ = Left(COTIMOV$(i%), 2)
          If COTIMOV1$ = "VM" Then
             If TRIM$(Control$("", "CODCONSU")) <> "" Then
                COTIMOV1$ = TRIM$(Control$("", "CODCONSU"))
             End If
          End If
          SQLX$ = SQLX$ + " OR CodiMovi = '" & COTIMOV1$ & "'  "
        End If
      End If
    Next i%
    If M1% > 0 Then SQLX$ = SQLX$ + " )"
    '
    SQLCODMOV$ = SQLX$
    If TRIM$(SQLCODMOV$) = "" Then
      Call COMUNI("Falta Seleccionar Tipo de Movimiento")
    End If

    Campos$ = "Año/A5;Mes/a0;Mes/A14;Código/A20;Dep./F5;T.Mov./A8;M./A4;C.Sal./F9.1;C.Ing./F9.1;Importe/f12.2,C1/F0;MONVAL/F0"

    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    FRMZELECHO.Caption = "ESTADÍSTICO DE CONSUMOS FORMATO MENSUAL"
    DEP% = 0: If XVALO(Text6) > 0 Then DEP% = XVALO(Text6)
    TMAT% = 0: If XVALO(Text3) > 0 Then TMAT% = XVALO(Text3)
    Dim obj As New RECORXET
    Set RS1 = obj.RS_ESTADISTICA_CONSUMO_AGRUPADA_X_MES(Des%, Has%, DEP%, SQLCODMOV$, TMAT%, C1%, C2%)
    Call Carga_MSF_Recordset(RS1, Campos$, FRMZELECHO.msf2, 1)
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    
    FRMZELECHO.mnuImprimir.Visible = True
    FRMZELECHO.Width = 13500
    FRMZELECHO.Show 1
    If FRMZELECHO.IMPRIMIR > 0 Then
     Dim MSF As msFlexGrid
     Set MSF = FRMZELECHO.msf2
     If FRMZELECHO.MSF1.Rows > 1 Then Set MSF = FRMZELECHO.MSF1
     Call CARGA_TABLA_IMPRE(MSF, Campos$, "ESTADMES", "ESTADÍSTICO DE CONSUMOS FORMATO MENSUAL", "TABLA_IMPRE")
     Call FINAL("?ESTADMES")
    End If

    

End Sub


Private Sub mnuEstadisticasSalidasMensual_Click()
    If TRIM(Text4) = "" Or TRIM(Text5) = "" Then
      Call COMUNI("Rango de Fecha no valido")
     
      Exit Sub
    End If
    
    Des% = FECHANUM(TRIM(Text4))
    Has% = FECHANUM(TRIM(Text5))
    
    If Des% > Has% Then
      Call COMUNI("Rango de Fecha no valido")
      Exit Sub
    End If
    '
    If TRIM(Text1) = "" Or TRIM(Text2) = "" Then
      Call COMUNI("Código no Valido.")
      Command12.Enabled = True
      Exit Sub
    End If
    C1% = CODINT(Text1)
    C2% = CODINT(Text2)
    If C1% <= 0 Or C2% <= 0 Then
      Call COMUNI("Código no Valido.")
      
      Exit Sub
    End If

    If IsFormLoaded(FRMZELECHO) Then
        ESTAENMEMORIA = 1
    Else
        Call HookForm(FRMZELECHO.msf2) 'estaba tirando un error cuando se cierra el formulario con doble clic
        Call HookForm(FRMZELECHO.MSF1) 'probablemente por que hace un hide, SI NO ESTA EN MEMORIA SE APLICA LA FUNCION.
    End If
    '
    ReDim COTIMOV$(List2.ListCount - 1)
    For i% = 0 To List2.ListCount - 1
      COTIMOV$(i%) = ""
    Next i%
    
    For i% = 0 To List2.ListCount - 1
      If List2.Selected(i%) = True Then
        HAYCOMOV% = 1
        COTIMOV$(i%) = List2.List(i%)
      End If
    Next i%
    
    If HAYCOMOV% = 0 Then
      Call COMUNI("Código de Movimiento no Valido.")
      Command12.Enabled = True
      Exit Sub
    End If
    M1% = 0
    For i% = 0 To List2.ListCount - 1 '8\\\OT-06-0218-WAR-05/06/06///
      If COTIMOV$(i%) <> "" Then
        If M1% = 0 Then
          COTIMOV1$ = Left(COTIMOV$(i%), 2)
          SQLX$ = SQLX$ + " AND ( CodiMovi = '" & COTIMOV1$ & "'"
          M1% = 1
        Else
          COTIMOV1$ = Left(COTIMOV$(i%), 2)
          If COTIMOV1$ = "VM" Then
             If TRIM$(Control$("", "CODCONSU")) <> "" Then
                COTIMOV1$ = TRIM$(Control$("", "CODCONSU"))
             End If
          End If
          SQLX$ = SQLX$ + " OR CodiMovi = '" & COTIMOV1$ & "'  "
        End If
      End If
    Next i%
    If M1% > 0 Then SQLX$ = SQLX$ + " )"
    '
    SQLCODMOV$ = SQLX$
    If TRIM$(SQLCODMOV$) = "" Then
      Call COMUNI("Falta Seleccionar Tipo de Movimiento")
    End If

    Campos$ = "Año/A5;Mes/a0;Mes/A14;Código/A20;Dep./F5;T.Mov./A8;M./A4;C.Sal./F9.1;C.Ing./F0;Importe/f12.2,C1/F0;MONVAL/F0"
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    FRMZELECHO.Caption = "ESTADÍSTICO DE CONSUMOS FORMATO MENSUAL SOLO MOV. DE SALIDA"
    DEP% = 0: If XVALO(Text6) > 0 Then DEP% = XVALO(Text6)
    TMAT% = 0: If XVALO(Text3) > 0 Then TMAT% = XVALO(Text3)
    Dim obj As New RECORXET
    Set RS1 = obj.RS_ESTADISTICA_CONSUMO_AGRUPADA_X_MES(Des%, Has%, DEP%, SQLCODMOV$, TMAT%, C1%, C2%, 1)
    Call Carga_MSF_Recordset(RS1, Campos$, FRMZELECHO.msf2, 1)
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    
    FRMZELECHO.mnuImprimir.Visible = True
    FRMZELECHO.Width = 13500
    FRMZELECHO.Show 1
    If FRMZELECHO.IMPRIMIR > 0 Then
     Dim MSF As msFlexGrid
     Set MSF = FRMZELECHO.msf2
     If FRMZELECHO.MSF1.Rows > 1 Then Set MSF = FRMZELECHO.MSF1
     Call CARGA_TABLA_IMPRE(MSF, Campos$, "ESTADMES", "ESTADÍSTICO DE CONSUMOS FORMATO MENSUAL", "TABLA_IMPRE")
     Call FINAL("?ESTADMES")
    End If

End Sub


Private Sub mnuInvertir_Click()
   TPI = List2.TopIndex
   For U& = 1 To List2.ListCount
     If List2.Selected(U& - 1) = True Then
         List2.Selected(U& - 1) = False
       Else
         List2.Selected(U& - 1) = True
     End If
   Next U&
   On Error Resume Next
   List2.TopIndex = TPI
   On Error GoTo 0
   DoEvents

End Sub

Private Sub mnusalir_Click()
   Call Command10_Click
End Sub

Private Sub Text1_Change()
    CI1% = CODINT%(Text1.TEXT)
    If CI1% > 0 Then
         Label1.Caption = DESCRIT$(CI1%)
       Else
         Label1.Caption = ""
    End If
End Sub

Private Sub Text2_Change()
    CI2% = CODINT%(Text2.TEXT)
    If CI2% > 0 Then
         Label2.Caption = DESCRIT$(CI2%)
       Else
         Label2.Caption = ""
    End If
End Sub

Private Sub Text3_Change()
    If XVALO(Text3.TEXT) < 256 Then
        VENDE% = XVALO(Text3.TEXT)
        VENDX$ = ECOARCH$("TATIMAT", Chr$(VENDE%))
        If VENDX <> "" Then
            Label4.Caption = VENDX$
            Exit Sub
          Else
            Label4.Caption = "Todos - Sin Clasificación"
        End If
    End If
End Sub

Private Sub Text6_Change()
    If XVALO(Text6.TEXT) < 256 Then
        VENDE% = XVALO(Text6.TEXT)
        VENDX$ = ECOARCH$("TADEPOSI", Chr$(VENDE%))
        If VENDX <> "" Then
            Label8.Caption = VENDX$
            Exit Sub
          Else
            Label8.Caption = "Todos - Sin Clasificación"
        End If
    End If
End Sub
Sub MATSINMOV()
    '
    If Des% < 1 Or Has% < Des% Then
       FEX = HOY(HO$)
       Has% = FEX
       HASS$ = FECHATEX$(FEX)
       DESS$ = "01" + Mid$(HASS$, 3)
       Des% = FECHANUM(DESS$)
    End If
    VDEF$ = MKI$(Des%) + MKI$(Has%)
    '
    OBX$ = OBJPLA$("DESHASFECHMAT", VDEF$)
    If OBX$ = "" Then
        Des% = 0: Has% = 0: PERIO$ = ""
        Exit Sub
    End If
    '
    Des% = CVI(Left$(OBX$, 2))
    Has% = CVI(Mid$(OBX$, 3, 2))
    If Des% > Has% Then Exit Sub
    FEX = Des%: PERIO$ = FECHATEX$(FEX)
    FEX = Has%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    DES1 = FETEXCOR1$(Des%)
    HAS1 = FETEXCOR1$(Has%)
    '
    TIPOMAT% = Asc(Mid$(OBX$, 5, 1))
    If ECOARCH$("TATIMAT", Chr$(TIPOMAT%)) = "" Then TIPOMAT% = 0
    '
    Call ABRESTOCKS
    '
    FIN& = NUMAS%
    i& = 0
    J& = 0
    For J& = 1 To NUMAS%
           Call TRACE(20, J&, FIN&)
           CI% = J&
           TIPOMATCOD% = TIMATE(CI%)
           If CI% > 0 And STATITEM%(CI%) > 0 And (TIPOMAT% = TIPOMATCOD% Or TIPOMAT% = 0) Then
              CONDI$ = "CANTSALID > 0.05 AND FECHMOV >= '" & DES1 & "' AND  FECHMOV <= '" & HAS1 & "' AND COD_INTE = " & CI%
              CAREG& = CantRegistros("MOVISTO", CONDI$)
              If CAREG& < 1 Then
                  i& = i& + 1
                  Call REPLA(TX$, MKI$(CI%), 1, 2)
                    ST# = STOCKACT(CI%)
                  Call REPLA(TX$, MKS$(ST#), 3, 4)
                    COSTOU# = COSUNI(CI%)
                  Call REPLA(TX$, MKS$(COSTOU#), 7, 4)
                    TOTALIT# = ST# * COSTOU#
                  Call REPLA(TX$, MKS$(TOTALIT#), 11, 4)
                  Call GRAREG("MASINMO", TX$, i&)
              End If
           End If
    Next J&
    Call SETULTREG("MASINMO", i&)
    Call FILESORT("MASINMO", "MASINMO", 1, 1, "ASC")
    Call HEADERS("MASINMO", "")
    Call HEADERS("MASINMO", "Desde: " & FECHATEX$(Des%) & " Hasta: " & FECHATEX$(Has%))
    If TIPOMAT% > 0 Then Call HEADERS("MASINMO", "Tipo Material: " & CStr(TIPOMAT%) & " - " & ECOARCH$("TATIMAT", MKI$(TIPOMAT%)))
    Call FINAL("*MASINMO")
    Call CIERRARCH("MASINMO")
 
    '
End Sub


Sub MATSINMOV_CONEXPLOSION()
    '
    CI1% = PRICOD%
    CI2% = ULTCOD%
    VDEF$ = MKI$(CI1%) + MKI$(CI2%) + Chr$(0)
    '
10  OBX$ = OBJPLA$("DESHASCODEP", VDEF$)
    If Len(OBX$) < 5 Then
        Exit Sub
    End If
    COD1$ = CODEXT$(CVI(Left$(OBX$, 2)))
    COD2$ = CODEXT$(CVI(Mid$(OBX$, 3, 2)))
    If TRIM$(COD1$) = "" Or TRIM$(COD2$) = "" Or COD2$ < COD1$ Then
        GoTo 10
    End If
    
    SUCUI% = Asc(Mid$(OBX$, 5, 1))
    If ECOARCH("ECODEP", Chr$(SUCUI%)) = "" Then
        SUCUI% = 0
    End If
    '
    Call DESHASFECHA(Des%, Has%, PERIO$)
    If PERIO$ = "" Then
      Exit Sub
    End If
    '
    DES1 = FETEXCOR1$(Des%)
    HAS1 = FETEXCOR1$(Has%)
    '
    Call ABRESTOCKS
    '
    FIECO% = FILENBR%("ASINMOEX")
    Dim ECM As String * 16
    '
    Z00$ = REGBLAN$("ASINMOEX")
    '
    Call HEADERS("ASINMOEX", "")
    Call HEADERS("ASINMOEX", "Desde Codigo: " + TRIM$(COD1$) + " - " + DESCRIT$(CI1%))
    Call HEADERS("ASINMOEX", "Hasta Codigo: " + TRIM$(COD2$) + " - " + DESCRIT$(CI2%))
    Call HEADERS("ASINMOEX", "Desde Fecha: " + FECHATEX$(Des%) + " - Hasta: " + FECHATEX$(Has%))
    If SUCUI% > 0 Then
        Call HEADERS("ASINMOEX", "Deposito: " + CStr(SUCUI%) + " - " + ECOARCH("ECODEP", Chr$(SUCUI%)))
    End If
    '
    Screen.MousePointer = 11
    FIN& = NUMAS%
    J& = 0
    DEPO% = XVALO(Text6)
    '
    For i& = 1 To FIN&
      Call TRACE(20, i&, FIN&)
      CI% = i&
      If CI% > 0 And STATITEM%(CI%) > 0 Then
        CONDI$ = "CANTSALID > 0.05 AND FECHMOV >= '" & DES1 & "' AND  FECHMOV <= '" & HAS1 & "' AND COD_INTE = " & CI%
        If DEPO% > 0 Then CONDI$ = CONDI$ & " AND DEPOMOVI=" & CStr(DEPO%)
        CAREG& = CantRegistros("MOVISTO", CONDI$)
        If CAREG& < 1 Then
          ACTIX% = 0
          CODYY$ = CODEXT$(CI%)
          If CODYY$ >= COD1$ Then
            If CODYY$ <= COD2$ Then
              PORMILLAR% = 0
              Call REPLA(Z$, MKI$(CI%), 1, 2)
              Call REPLA(Z$, MKI$(CI%), 3, 2)
              Call REPLA(Z$, UM$, 5, 2)
              Call REPLA(Z$, MKS$(0), 7, 4)
                  STACTUAL# = STOCKACT(CI%)
              Call REPLA(Z$, MKS$(STACTUAL#), 11, 4)
              '
              J& = J& + 1
              Call GRAREG("ASINMOEX", Z$, J&)
              '
              Call LEEEXPLO(CODYY$, 1)
              If CAIT% > 0 Then
                ' DEJA UN RENGLON EN BLANCO
                Z$ = Z00$
                Call REPLA(Z$, MKI$(CI%), 1, 2)
                J& = J& + 1
                Call GRAREG("ASINMOEX", Z$, J&)
                
                For k% = 1 To CAIT%
                  CI1% = CIJ%(k%)
                  CAN = USOI(k%)
                  PD% = PDI%(k%)
                  If CI1% > 0 Then
                    If CI1% <= NUMAS% Then
                      UM$ = Unimed$(CI1%)
                      Call CONVERUNID(CAN, UM$)
                      If Abs(CAN) > 0.00005 Then
                        Z$ = Z00$
                        Call REPLA(Z$, MKI$(CI%), 1, 2)
                        Call REPLA(Z$, MKI$(CI1%), 3, 2)
                        Call REPLA(Z$, UM$, 5, 2)
                        Call REPLA(Z$, MKS$(CAN), 7, 4)
                            STACTUAL# = STOCKACT(CI1%)
                        Call REPLA(Z$, MKS$(STACTUAL#), 11, 4)
                        '
                        J& = J& + 1
                        Call GRAREG("ASINMOEX", Z$, J&)
                        ACTIX% = 1
                      End If
                    End If
                  End If
                Next k%
              End If
            End If
          End If
          '
          'If ACTIX% > 0 Then
          '  Z$ = Z00$
          '  Call REPLA(Z$, MKI$(CI%), 1, 2)
          '  j& = j& + 1
          '  Call GRAREG("ASINMOEX", Z$, j&)
          'End If
          '
        End If
      End If
      '
    Next i&
    '
    Screen.MousePointer = 1
    Call SETULTREG("ASINMOEX", J&)
    Call CIERRARCH("ASINMOEX")
    Call FILESORT("ASINMOEX", "", 1, 1, "ASC")
    '
    Call FINAL("*ASINMOEX")
    
End Sub
