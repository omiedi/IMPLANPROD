VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form PREPAP07 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Preparación de Pedidos de Clientes"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5850
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
   ScaleHeight     =   7110
   ScaleWidth      =   5850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   105
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   34
      Top             =   735
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   7260
      Left            =   4935
      ScaleHeight     =   7230
      ScaleWidth      =   1455
      TabIndex        =   27
      Top             =   0
      Width           =   1485
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   150
         ScaleHeight     =   165
         ScaleWidth      =   570
         TabIndex        =   37
         Top             =   4320
         Width           =   600
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   38
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton AGenRep 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   150
         Picture         =   "PREPAP07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Acceso a Consultas Pre-Configuradas"
         Top             =   3150
         Width           =   600
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
         Height          =   600
         Left            =   120
         Picture         =   "PREPAP07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Control Depósito de Expedición"
         Top             =   2310
         Width           =   600
      End
      Begin VB.CommandButton Command26 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   150
         Picture         =   "PREPAP07.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Configurar Etiqueta de Expedición"
         Top             =   1575
         Width           =   600
      End
      Begin VB.CommandButton command1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   150
         Picture         =   "PREPAP07.frx":0A56
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Cierra y Abandona la Aplicación"
         Top             =   315
         Width           =   600
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
         Height          =   600
         Left            =   150
         Picture         =   "PREPAP07.frx":0BA0
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   945
         Width           =   600
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00CDDADA&
      Caption         =   "CONSULTAS Y LISTADOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2190
      Left            =   150
      TabIndex        =   8
      Top             =   2520
      Width           =   4635
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   420
         ScaleHeight     =   615
         ScaleWidth      =   3765
         TabIndex        =   12
         Top             =   525
         Width           =   3795
         Begin VB.CommandButton Command64 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   3180
            Picture         =   "PREPAP07.frx":0EAA
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Preparados por Código de Producto Mostrando Total Pedido"
            Top             =   0
            Width           =   600
         End
         Begin VB.CommandButton Command62 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   2580
            Picture         =   "PREPAP07.frx":11B4
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Preparados por Cliente Mostrando Total Pedido"
            Top             =   0
            Width           =   600
         End
         Begin VB.CommandButton Command61 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   0
            Picture         =   "PREPAP07.frx":15F6
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Detalle de Pedidos Preparados por Número"
            Top             =   0
            Width           =   560
         End
         Begin VB.CommandButton Command63 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   1980
            Picture         =   "PREPAP07.frx":1900
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Preparados por Número de Pedido Mostrando Total Pedido"
            Top             =   0
            Width           =   600
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "B.D. Pedidos Preparados"
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
            Left            =   735
            TabIndex        =   17
            Top             =   105
            Width           =   1485
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00CDDADA&
         Caption         =   "DISPOSITIVO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   610
         Left            =   420
         TabIndex        =   9
         Top             =   1365
         Width           =   3795
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Pantalla"
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
            Height          =   225
            Left            =   525
            TabIndex        =   11
            Top             =   270
            Value           =   -1  'True
            Width           =   1065
         End
         Begin VB.OptionButton Option2 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Impresión"
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
            Height          =   225
            Left            =   2205
            TabIndex        =   10
            Top             =   270
            Width           =   1170
         End
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00CDDADA&
      Caption         =   "TRANSACCIONES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2265
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   4635
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Index           =   3
         Left            =   1465
         ScaleHeight     =   615
         ScaleWidth      =   2400
         TabIndex        =   24
         Top             =   1440
         Width           =   2430
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
            Height          =   645
            Left            =   0
            Picture         =   "PREPAP07.frx":1C0A
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Imprimir Etiquetas de Despacho"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Imprimir Etiquetas de Expedición"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   735
            TabIndex        =   25
            Top             =   105
            Width           =   1635
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Index           =   2
         Left            =   1100
         ScaleHeight     =   615
         ScaleWidth      =   2400
         TabIndex        =   21
         Top             =   960
         Width           =   2430
         Begin VB.CommandButton Command45 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   0
            Picture         =   "PREPAP07.frx":1D54
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Cancelar Preparación de Pedidos (Vuelta a Almacen)"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Cancelar la Preparación"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   840
            TabIndex        =   22
            Top             =   45
            Width           =   1425
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Index           =   0
         Left            =   735
         ScaleHeight     =   615
         ScaleWidth      =   2400
         TabIndex        =   18
         Top             =   480
         Width           =   2430
         Begin VB.CommandButton Command46 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   0
            Picture         =   "PREPAP07.frx":205E
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Preparación de Pedidos"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Preparación de Pedidos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   840
            TabIndex        =   19
            Top             =   45
            Width           =   1215
         End
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Listados de Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2640
         Left            =   2625
         TabIndex        =   1
         Top             =   4830
         Width           =   4005
         Begin VB.CommandButton Command16 
            Caption         =   "Análisis A-B-C"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2205
            TabIndex        =   7
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command17 
            Caption         =   "Estadísticas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   210
            TabIndex        =   6
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command20 
            Caption         =   "Críticos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2205
            TabIndex        =   5
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command21 
            Caption         =   "Consignación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   210
            TabIndex        =   4
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command22 
            Caption         =   "Negativos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2205
            TabIndex        =   3
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command23 
            Caption         =   "Inventarios"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   210
            TabIndex        =   2
            Top             =   420
            Width           =   1590
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Controles no Visibles"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1905
      Left            =   150
      TabIndex        =   30
      Top             =   5040
      Visible         =   0   'False
      Width           =   4635
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
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
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   210
         TabIndex        =   31
         Top             =   315
         Width           =   1065
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "PREPAP07.frx":2928
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "PREPAP07.frx":2B5C
      TabIndex        =   35
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
End
Attribute VB_Name = "PREPAP07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub AGenRep_Click()
    '
    AGenRep.Enabled = False
    Call ACONSUL_PRE(TRIM$(UCase$(Name)))
    AGenRep.Enabled = True
    '
End Sub

Private Sub AGenRep_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   '
   If Button = 2 Then
      Call CONECTA_CONSUL_PRE(TRIM$(UCase$(Name)))
   End If
   '
End Sub

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub


Private Sub Command2_Click()
   Command2.Enabled = False
   Call MENSERR(24, "No Requerido ni Disponible\en la Presente Versión.")
   GoTo 99
   '
   '
   '
   If DERACCE%("PREPAPED", "Preparar Pedidos de Clientes") < 2 Then
     GoTo 99
   End If
   '
   Dim CAPRENET(32767), CASTOEXP(32767)
   DEPEX% = Int(Val(Control$("", "DEPOEXPE")))
   DETER% = Int(Val(Control$("", "DEPOTERM")))
   If DEPEX% < 1 Or DETER% < 1 Or DEPEX% = DETER% Then
     Call MENSERR(24, "Proceso Imposible.\  \Falta Configurar Depósitos\de Expedición\y Productos Terminados.")
     GoTo 99
   End If
   '
   Screen.MousePointer = 11
   RFF$ = RECFILT$("DESPAPED", 7, Chr$(1))
   FIN& = Len(RFF$)
   '
   Screen.MousePointer = 11
   For U& = 1 To FIN& Step 4
     Call TRACE(20, U&, FIN&)
     REDE& = CVS(Mid$(RFF$, U&, 4))
     ' AQUI VERIFICAR ESTADO DE PREPARACION DEL REGISTRO Y PASARLO A STATUS 3 SI ESTA CUMPLIDO
'''     Call VERIDESPRE(REDE&, HUBORECU%)
     'If STADESPE%(REDE&) > 1 Then GoTo 29
   Next U&
   '
   RFF$ = RECFILT$("DESPAPED", 7, Chr$(11))
   NPXA& = 0: JJ& = 0: FIN& = Len(RFF$)
   For U& = 1 To FIN& Step 4
     Call TRACE(20, U&, FIN&)
     REDE& = CVS(Mid$(RFF$, U&, 4))
     XX$ = REGLEIDO$("DESPAPED", REDE&)
     '
     If Mid$(XX$, 51, 1) <> "R" Then   ' ES UNA REPARACIóN
       NPX& = CVS(Left$(XX$, 4))
       CI% = CVI(Mid$(XX$, 33, 2))
       CDXX$ = TRIM$(CODEXT$(CI%))
       CAPEDIT = CVS(Mid$(XX$, 75, 4))
       ACUPRE = CVS(Mid$(XX$, 79, 4))
       ACUDES = CVS(Mid$(XX$, 95, 4))
       '
       If ACUDES > ACUPRE + 0.005 Then
         Call MENSERR(24, "Pedido " + TRIM$(Str$(NPX&)) + " - Cod: " + CDXX$ + "\Acumulado Despachado Supera Acumulado Preparado.")
         GoSub BORRAMOPRE
       End If
       '
       If ACUDES > CAPEDIT + 0.005 Then
         Call MENSERR(24, "Pedido " + TRIM$(Str$(NPX&)) + " - Cod: " + CDXX$ + "\Acumulado Despachado Supera Total Pedido.")
         GoSub BORRAMOPRE
       End If
       '
       If ACUPRE > CAPEDIT + 0.005 Then
         Call MENSERR(24, "Pedido " + TRIM$(Str$(NPX&)) + " - Cod: " + CDXX$ + "\Acumulado Preparado Supera Total Pedido.")
         GoSub BORRAMOPRE
       End If
       '
       '
       CAVERI$ = Mid$(XX$, 113, 10)
       CAVERJ$ = MKS$(Val(Left$(CAVERI$, 6))) + MKI$(Val(Mid$(CAVERI$, 8, 3)))
       '
       RGG$ = RECFILT$("MOVISTO", 17, CAVERJ$) ' SELECCIONA MOVIMIENTOS DEL ITEM
       FIN& = Len(RGG$)
       For UH& = 1 To FIN& Step 4
         REMOX& = CVS(Mid$(RGG$, UH&, 4))
         MM$ = REGLEIDO$("MOVISTO", REMOX&)
         
         If CVI(Mid$(MM$, 3, 2)) <> CI% Then
           GoSub BORRAMOPRE
         End If
       Next UH&
       '
       CAPRENET(CI%) = CAPRENET(CI%) + ACUPRE - ACUDES
     End If
19 Next U&
   '
   NMS% = NUMAS%
   FIN& = ULTREG("MOVISTO")
   REBLA$ = REGBLAN$("MALSTOXP")
   JXP& = 0
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     MVS$ = REGLEIDO$("MOVISTO", U&)
     DEP% = Asc(Mid$(MVS$, 83, 1))
     If DEP% = DEPEX% Then
       CI% = CVI(Mid$(MVS$, 3, 2))
       If CI% > 0 Then
         If CI% <= NMS% Then
           CASTOEXP(CI%) = CASTOEXP(CI%) + CVD(Mid$(MVS$, 88, 8)) - CVD(Mid$(MVS$, 96, 8))
         End If
       End If
     End If
   Next U&
   '
   For CI% = 1 To NUMAS%
     If Abs(CASTOEXP(CI%) - CAPRENET(CI%)) > 0.05 Then
       YX$ = REBLA$
       Call REPLA(YX$, MKI$(CI%), 1, 2)
       Call REPLA(YX$, MKS$(CAPRENET(CI%)), 3, 4)
       Call REPLA(YX$, MKS$(CASTOEXP(CI%)), 7, 4)
       Call REPLA(YX$, MKS$(CAPRENET(CI%) - CASTOEXP(CI%)), 11, 4)
       JXP& = JXP& + 1
       Call GRAREG("MALSTOXP", YX$, JXP&)
     End If
   Next CI%
   '
   Call SETULTREG("MALSTOXP", JXP&)
   Call CIERRARCH("MALSTOXP")
   Screen.MousePointer = 1
   '
   If JXP& > 0 Then
     If COMALTER%("Se Han Detectado Diferencias\entre los Preparados\y el Stock del Depósito de Expedición.\   \¿ Desea Emitir un Listado de Control ?\\Si, Listar\No, Continuar") = 1 Then
       Call FILESORT("MALSTOXP", "", 1, 1, "ASC")
       Call FINAL("*MALSTOXP")
       GoTo 99
     End If
   End If
   '
   Call COMUNI("Fin del Control")
   '
99 Command2.Enabled = True
Exit Sub
   '
BORRAMOPRE:
   '
   NPXX& = NPX&
   TTXX$ = "Se ha Detectado una Inconsistencia"
   TTXX$ = TTXX$ + "\en la Preparación del Pedido " + TRIM$(Str$(NPXX&))
   TTXX$ = TTXX$ + "\  \Deben Revertirse Todos los Movimientos"
   TTXX$ = TTXX$ + "\de Preparación y Despacho de este Pedido"
   If COMALTER%(TTXX$ + "\\Cancelar\Aceptar") = 2 Then
     REBLA$ = REGBLAN$("MOVISTO")
     ABUSBLA$ = "VP." + TRIM$(Str$(NPXX&))
     ABUSBLB$ = "DM." + TRIM$(Str$(NPXX&))
     ABUSBLC$ = "PL." + TRIM$(Str$(NPXX&))
     FIN& = ULTREG&("MOVISTO")
     For KKU& = 1 To FIN
       Call TRACE(20, KKU&, FIN&)
       XX$ = REGLEIDO$("MOVISTO", KKU&)
       ABUSX$ = TRIM$(Mid$(XX$, 23, 10))
       If ABUSX$ = ABUSBLA$ Or ABUSX$ = ABUSBLB$ Then
           Call GRAREG("MOVISTO", REBLA$, KKU&)
         ElseIf ABUSX$ = ABUSBLC$ Then
           If CVI(Left$(XX$, 2)) < FECHANUM("07/11/05") Then
             If CVD(Mid$(XX$, 96, 8)) >= 0.005 Then
               Call GRAREG("MOVISTO", REBLA$, KKU&)
             End If
           End If
       End If
     Next KKU&
     Call CIERRARCH("MOVISTO")
     '
     RGH$ = RECFILT$("DESPAPED", 1, MKS$(NPX&))
     For KU& = 1 To Len(RGH$) Step 4
       REDEY& = CVS(Mid$(RGH$, KU&, 4))
       XYX$ = REGLEIDO$("DESPAPED", REDEY&)
       Call REPLA(XYX$, Chr$(1), 27, 1)
       Call GRAREG("DESPAPED", XYX$, REDEY&)
     Next KU&
     Call CIERRARCH("DESPAPED")
     '
'     TTXYZ$ = Space$(76)
''     HOII% = HOY(HOS$)
'     Call REPLA(TTXYZ$, HOS$, 1, 8)
'     Call REPLA(TTXYZ$, "Cancelacion General Preparados y Despacho ", 11, 56)
'     Call REPLA(TTXYZ$, USERTER$, 69, 24)
'     Call REGAPP("ANOTAPED", MKS$(NPXX&) + TTXYZ$)
'     TTXYZ$ = Space$(76)
'     Call REPLA(TTXYZ$, "por Aparente Sobre-Embarque. ", 11, 56)
'     Call REGAPP("ANOTAPED", MKS$(NPXX&) + TTXYZ$)
'     Call REGAPP("ANOTAPED", MKS$(NPXX&) + Space$(88))
'     Call REGAPP("ANOTAPED", MKS$(NPXX&) + String$(88, "="))
'     Call CIERRARCH("ANOTAPED")
'     Call FRESHECHO("ANOTAPED")
     '
      '\\\OT-08-0769-OD-29/08/08///

      '
      TTXYZ$ = Space$(76)
      HOII% = HOY(HOS$)

      Call REPLA(TTXYZ$, HOS$, 1, 8)
      Call REPLA(TTXYZ$, "Cancelacion General Preparados y Despacho ", 11, 56)
      Call REPLA(TTXYZ$, USERTER$, 69, 24)
      Call REPLA(TTX11$, "por Aparente Sobre-Embarque. ", 11, 56)
      strSQL = "Select * from Pedenca where NroPed=" & NPX&
      Dim PedEnca As ADODB.Recordset
      Set PedEnca = New ADODB.Recordset
      PedEnca.Open FILTSQL$(strSQL), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
      ANOTANTE$ = ""
      '
      With PedEnca
        If Not (.EOF And .BOF) Then
            ANOTA$ = TTXYZ$
            ANOTA$ = ANOTA$ + Space$(88) + Chr$(10) + Chr$(13)
            ANOTA$ = ANOTA$ + "por Aparente Sobre-Embarque.  " + Chr$(10) + Chr$(13)
            ANOTA$ = ANOTA$ + TTX11$ + Chr$(10) + Chr$(13)
            ANOTA$ = ANOTA$ + String$(88, "=") + Chr$(10) + Chr$(13)
            ANOTANTE$ = !AnotaPed
            !AnotaPed = ANOTANTE$ + ANOTA$
            .Update
        End If
      End With
      PedEnca.Close
      Set PedEnca = Nothing
      '
      '\\\OT-08-0769-OD-FIN///

      Call COMUNI("Vuelva a Ingresar\a Preparacion de Pedidos")
      Call FINAL("")

   End If
Return
'
End Sub

'
Private Sub Command26_Click()
   Command26.Enabled = False
   If DERACCE%("STPPREP", "Configuración de Funcionamiento de Preparaciones") = 2 Then
       CFGDES07.Show 1
   End If
   Command26.Enabled = True
End Sub

Private Sub Command43_Click()
    Command43.Enabled = False
    Screen.MousePointer = 11
    Dim ITPEDI$(), CAPENDE(), NUPEDI&()
    MAXPED& = 4096
    ReDim ITPEDI$(MAXPED&), CAPENDE(MAXPED&), NUPEDI&(MAXPED&)
    '
    DEPX% = DEPOEXPE%
    Call COPYSTRU("INDIPEDI", "IPEDPREP")
    
    SQLX$ = "SELECT DISTINCT PEDDETA.NroPed, PEDDETA.VENREP, PEDENCA.Referencia as REFEREN FROM PEDDETA WITH(NOLOCK)"
    SQLX$ = SQLX$ + " INNER JOIN PEDENCA"
    SQLX$ = SQLX$ + " ON PEDDETA.NroPed = PEDENCA.NroPed"
    SQLX$ = SQLX$ + " WHERE PEDDETA.NroPed > 0 "
    SQLX$ = SQLX$ + " AND PEDENCA.Status < 9 "                ' NO ANULADOS
    SQLX$ = SQLX$ + " AND PEDDETA.Status <> 8 "               ' NO ANULADOS
    SQLX$ = SQLX$ + " AND PEDDETA.Status <> 9 "               ' NO ANULADOS
    SQLX$ = SQLX$ + " AND PEDENCA.CODIEMPR = " & CodiEmp% & " "
    SQLX$ = SQLX$ + " AND PEDDETA.CODIEMPR = " & CodiEmp% & " "
    SQLX$ = SQLX$ + " AND AcuPrepa > AcuDespa"
    'SQLX$ = SQLX$ + " AND PEDDETA.VenRep = 'V' "             ' Excluye Reparaciones
    SQLX$ = SQLX$ + " ORDER BY PEDDETA.NroPed ASC"
    Set IPEPE = New ADODB.Recordset
    IPEPE.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    '
    JJ& = 0
    With IPEPE
       On Error Resume Next
       .MoveLast
       If Not Err Then
          On Error GoTo 0
          FIN& = .RecordCount
          .MoveFirst
          Do While Not .EOF
            YY$ = MKS$(!NROPED) + TRIM$(!Referen)
            REPAR$ = "": If !VenRep = "R" Then REPAR$ = " (Reparación)"
            YY$ = YY$ + REPAR$
            JJ& = JJ& + 1
            Call TRACE(20, JJ&, FIN&)
            Call GRAREG("IPEDPREP", YY$, JJ&)
          .MoveNext
          Loop
       End If
    End With
    '
    Call SETULTREG("IPEDPREP", JJ&)
    Call FRESHECHO("IPEDPREP")
    If JJ& < 1 Then
      Call COMUNI("No hay Pedidos Preparados\Pendientes de Despacho.")
      GoTo 99
    End If
    '
3   Screen.MousePointer = 1
    Call SELECHO("IPEDPREP/Indice de Pedidos Preparados")
    NPX& = Val(VALACT1$("IPEDPREP"))
    If NPX& > 0 Then
      Call ETIDESPA(NPX&)
      GoTo 3
    End If
    '
99  Command43.Enabled = True
End Sub

Private Sub Command45_Click()
   '
   Static DEPEX%, DETER%
   '
   Command45.Enabled = False
   If DERACCE%("PREPAPED", "Preparar Pedidos de Clientes") < 2 Then
      GoTo 99
   End If
   '
   Screen.MousePointer = 11
   If DEPEX% < 1 Or DETER% < 1 Then
     DEPEX% = Int(Val(Control$("", "DEPOEXPE")))
     DETER% = Int(Val(Control$("", "DEPOTERM")))
     If DEPEX% < 1 Or DETER% < 1 Or DEPEX% = DETER% Then
       Call MENSERR(24, "Proceso Imposible.\  \Falta Configurar Depósitos\de Expedición\y Productos Terminados.")
       GoTo 99
     End If
   End If
   '
   Screen.MousePointer = 11
   Call COPYSTRU("INDIPEDI", "IPEDPREP")
   SQLX$ = "SELECT DISTINCT PEDDETA.NroPed, PEDENCA.Referencia as REFEREN FROM PEDDETA WITH(NOLOCK)"
   SQLX$ = SQLX$ + " INNER JOIN PEDENCA"
   SQLX$ = SQLX$ + " ON PEDDETA.NroPed = PEDENCA.NroPed"
   SQLX$ = SQLX$ + " WHERE PEDDETA.NroPed > 0 "
   SQLX$ = SQLX$ + " AND PEDENCA.Status < 9 "                ' NO ANULADOS
   SQLX$ = SQLX$ + " AND PEDDETA.Status <> 8 "
   SQLX$ = SQLX$ + " AND PEDDETA.Status <> 9 "
   SQLX$ = SQLX$ + " AND PEDDETA.StaDespa < 2 "              ' elimina los cancelados
   SQLX$ = SQLX$ + " AND PEDENCA.CODIEMPR = " & CodiEmp% & " "
   SQLX$ = SQLX$ + " AND PEDDETA.CODIEMPR = " & CodiEmp% & " "
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
      If Not Err Then
         On Error GoTo 0
         FIN& = .RecordCount
         .MoveFirst
         Do While Not .EOF
           YY$ = MKS$(!NROPED) + !Referen
           JJ& = JJ& + 1
           Call TRACE(20, JJ&, FIN&)
           Call GRAREG("IPEDPREP", YY$, JJ&)
         .MoveNext
         Loop
      End If
   End With
   '
   Call SETULTREG("IPEDPREP", JJ&)
   Call FRESHECHO("IPEDPREP")
   If JJ& < 1 Then
     Call COMUNI("No hay Pedidos Preparados\Pendientes de Despacho.")
     GoTo 99
   End If
   '
3  Screen.MousePointer = 1
   Call SELECHO("IPEDPREP/Indice de Pedidos Pendientes de Despacho")
   NPX& = Val(VALACT1$("IPEDPREP"))
   If NPX& < 1 Then GoTo 99
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
   REBLA$ = REGBLAN$("DESPAPED")
   SQLX$ = " SELECT * FROM PEDDETA "
   SQLX$ = SQLX$ + " WHERE NroPed = " & NPX& & " "   ' COINCIDE NÚMERO DE PEDIDO
   SQLX$ = SQLX$ + " order by NuOrItem "
   Dim DESPAPE As ADODB.Recordset
   Set DESPAPE = New ADODB.Recordset
   DESPAPE.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
   '
   With DESPAPE
      On Error GoTo 0 'Resume Next
      .MoveLast
      If Err Then
         On Error GoTo 0
         Call MENSERR(24, "Imposible Des-Preparar Pedido " + TRIM$(Str$(NPX&)))
         GoTo 3
      End If
      On Error GoTo 0
      .MoveFirst
      Do While Not .EOF
         CI% = !CODIINT
         STATU% = !Status
         NCLIE% = !NROCLIE
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
           Call REPLA(Y$, MKD$(0), 163, 8)
           Call REPLA(Y$, MKS$(NORITL&), 189, 4)
           Call REGAPP("!BOLREDET", Y$)
         End If
      .MoveNext
      Loop
    End With
    '
12  VTB% = VENTABU%("DESPAPED/NC/TITUPAN=Cancelación de Preparación - " + TRIM$(rasocli$(NCLIE%)))
    If VTB% < 0 Or ULTREG&("!BOLREDET") < 1 Then
      GoTo 99
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
    Call MENSERR(24, "Transacción Abortada !!!\Faltan Cantidades o Códigos.")
    GoTo 3
    '
13  For YX& = 1 To ULTREG&("!BOLREDET")
      YY$ = REGLEIDO$("!BOLREDET", YX&)
      CI% = CVI(Mid$(YY$, 83, 2))
      CAPREP# = CVD(Mid$(YY$, 163, 8))
      MAXX# = CVD(Mid$(YY$, 89, 8))
      If CAPREP# > MAXX# + 0.05 Then
        Call MENSERR(24, "Imposible Des-Preparar Cantidad\Mayor a Pendiente de Despacho.")
        GoTo 12
      End If
    Next YX&
    '
15  CONTA% = ULTREG&("!BOLREDET")
    If CONTA% > 0 Then
      Call CIERRARCH("MOVISTO")
      Screen.MousePointer = 11  'Call BLIMESS("Registrando Recepcion ...")
      '
      TTXXX$ = "Materiales Vendidos"
      '
16    TTXYZ$ = Space$(76)
      HOII% = HOY(HOS$)
      ANOTA$ = ""
      Call REPLA(TTXYZ$, HOS$, 1, 8)
      Call REPLA(TTXYZ$, "Cancela Preparación de " + TTXXX$ + ":", 11, 56)
      Call REPLA(TTXYZ$, USERTER$, 69, 24)
      ANOTA$ = ANOTA$ + TTXYZ$ + Chr$(10) + Chr$(13)
      ANOTA$ = ANOTA$ + Space$(88) + Chr$(10) + Chr$(13)
      '
      AINDI% = 0
      For YX& = 1 To ULTREG&("!BOLREDET")
        YY$ = REGLEIDO$("!BOLREDET", YX&)
        CI% = CVI(Mid$(YY$, 83, 2))
        CAPREP# = CVD(Mid$(YY$, 163, 8))
        LOTE$ = Mid$(YY$, 69, 12)
        If CI% > 0 Then
          If CI% <= NUMAS% Then
            If CAPREP# > 0.005 Then
              CAPRES = CAPREP#
              NORII% = CVS(Mid$(YY$, 189, 4))
              ' ANOTAR EN ANOTAPED
              TTXX$ = Space$(76)
              CIXI% = CI%
              Call REPLA(TTXX$, "- " + CODEXT$(CIXI%), 11, 18)
              Call REPLA(TTXX$, DESCRIT0$(CIXI%), 29, 38)
              Call REPLA(TTXX$, CSTRING$(MKS$(CAPRES), 3, 10, 1, 2), 57, 10)
              ANOTA$ = ANOTA$ + Left$(TTXX$, 96) + Chr$(10) + Chr$(13)
              ' REALIZAR TRANSFERENCIA DE STOCK
              '
              CAVER$ = TRIM$(Str$(NPX&))
              While Len(CAVER$) > 6: CAVER$ = Mid$(CAVER$, 2): Wend
              While Len(CAVER$) < 6: CAVER$ = "0" + CAVER$: Wend
              IK$ = TRIM$(Str$(NORII%))
              CAVERI$ = CAVER$ + "-" + IK$
              FEMOVI% = HOY(HO$)
              DOPRI$ = "VP." & CStr(NPX&)
              DOSEC$ = CAVERI$
              REFMO$ = "Canc.Prep.O/V-" + CAVERI$
              CAPREPA = CAPREP#
              Call MOVISTOK(FEMOVI%, CI%, LOTE$, "PP", DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFMO$, 0, VUNI#, 1, NCLIE%, DEPEX%, (-1) * CAPREPA, NPX&, NORII%)
              Call MOVISTOK(FEMOVI%, CI%, LOTE$, "VP", DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFMO$, 0, VUNI#, 1, NCLIE%, DETER%, CAPREPA, NPX&, NORII%)
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
      PedEnca.Open FILTSQL$(strSQL), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
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
      Call VERIDESPRE(NPX&, 0, HUBORECU%)
      '
    End If
    '
    GoTo 3
    '
99  Screen.MousePointer = 1
    Command45.Enabled = True
    '
End Sub

Function STADESPE%(REGIST&)
    Static DPEX%
    If DPEX% < 1 Then
      DPEX% = DEPOEXPE%
    End If
    '
    XX$ = REGLEIDO$("DESPAPED", REGIST&)
    NPXX& = CVS(XX$)
    CII% = CVI(Mid$(XX$, 33, 2))
    CAPEDIT = CVS(Mid$(XX$, 75, 4))
    STATUDES% = Asc(Mid$(XX$, 27, 1))
    CAVERI$ = Mid$(XX$, 113, 10)
    '
    If Mid$(XX$, 51, 1) = "R" Then GoTo 90 ' PARA REPARACIONES EL STATUS ES EL MARCADO
    '
    If TRIM$(CAVERI$) = "" Then
      ' ESTO ES PARA RECUPERAR EL IDENTIFICADOR DE PEDIDO DAÑADO
      CAVER$ = TRIM$(Str$(NPXX&))
      While Len(CAVER$) > 6: CAVER$ = Mid$(CAVER$, 2): Wend
      While Len(CAVER$) < 6: CAVER$ = "0" + CAVER$: Wend
      IKITEM% = 0
      '
      RGH$ = RECFILT$("DESPAPED", 1, MKS$(NPXX&))
      For KI% = 1 To 999
        CAVERI$ = CAVER$ + "-" + TRIM$(Str$(KI%))
        For KLU& = 1 To Len(RGH$) Step 4
          KUU& = CVS(Mid$(RGH$, KLU&, 4))
          XZ$ = REGLEIDO$("DESPAPED", KUU&)
          If TRIM$(Mid$(XZ$, 113, 10)) = CAVERI$ Then GoTo 19
        Next KLU&
        Call REPLA(XX$, CAVERI$, 113, 10)
        Call GRAREG("DESPAPED", XX$, REGIST&)
        GoTo 20
19    Next KI%
    End If
    '
20  CAVERJ$ = MKS$(Val(Left$(CAVERI$, 6))) + MKI$(Val(Mid$(CAVERI$, 8, 3)))
    RFF$ = RECFILT$("MOVISTO", 17, CAVERJ$)
    CAYAPRE = 0
    For KU& = 1 To Len(RFF$) Step 4
      REMO& = CVS(Mid$(RFF$, KU&, 4))
      MM$ = REGLEIDO$("MOVISTO", REMO&)
      If Asc(Mid$(MM$, 83, 1)) = DPEX% Then      ' MOVIMIENTO EN DEPOSITO DE EXPEDICION
        If UCase$(Mid$(MM$, 21, 2)) = "PP" Then  ' MOVIMIENTO PREPARACION DE PEDIDO
          'If Mid$(MM$, 112, 6) = CAVERJ$ Then    ' COINCIDE PEDIDO E ITEM
          If CVI(Mid$(MM$, 3, 2)) = CII% Then      ' COINCIDE ARTICULO
            CAYAPRE = CAYAPRE + CVD(Mid$(MM$, 88, 8)) - CVD(Mid$(MM$, 96, 8))
          End If
        End If
      End If
    Next KU&
    Call REPLA(XX$, MKS$(CAYAPRE), 79, 4)
    ' SE SUPRIME LO QUE SIGUE PORQUE LA MARCA ES CUANDO TERMINA DESPACHO
    'If CAYAPRE >= CAPEDIT - 0.05 Then
    '  STATUDES% = 3   ' CUMPLIDO
    '  Call REPLA(XX$, Chr$(3), 27, 1)
    'End If
    Call GRAREG("DESPAPED", XX$, REGIST&)
    '
    ' SI CANT PREPARADA > QUE PEDIDA MARCA COMO YA PREPARADO
    If CAYAPRE >= CAPEDIT - 0.05 Then
      STATUDES% = 2
    End If
    '
90  STADESPE% = STATUDES%
    '
End Function
'
Private Sub Command46_Click()
   Command46.Enabled = False
   If DERACCE%("PREPAPED", "Preparar Pedidos de Clientes") < 2 Then
     GoTo 99
   End If
   '
   Dim CAYADES(32767), ACTISTO%(32767)
   DEPEX% = Int(Val(Control$("", "DEPOEXPE")))
   DETER% = Int(Val(Control$("", "DEPOTERM")))
   If DEPEX% < 1 Or DETER% < 1 Or DEPEX% = DETER% Then
     Call MENSERR(24, "Proceso Imposible.\  \Falta Configurar Depósitos\de Expedición\y Productos Terminados.")
     GoTo 99
   End If
   '
2  Screen.MousePointer = 11
   Call APERBASDAT("STOCKS")
   Call APERBASDAT("PEDCLI")
   Call COPYSTRU("INDIPEDI", "PEPENPRE")
   SQLX$ = "SELECT DISTINCT NroPed FROM PEDDETA WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE NroPed > 0 "
   SQLX$ = SQLX$ + " AND AcuPrepa IS NULL "
   SQLX$ = SQLX$ + " AND VenRep <> 'R' "             ' Excluye Reparaciones
   Dim IPEPE As ADODB.Recordset
   Set IPEPE = New ADODB.Recordset
'   IPEPE.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
   Set IPEPE = READSET(SQLX$)
   '
   With IPEPE
      On Error Resume Next
      FIN& = .RecordCount
      If FIN& > 0 Then
         On Error GoTo 0
         Call COMUNI("FLEXOFT Debe Ahora Revisar\la Situación de Preparación y Despacho.")
         Screen.MousePointer = 11
         .MoveFirst
         KU& = 0
         Do While Not .EOF
            On Error GoTo 0
            KU& = KU& + 1
            Call TRACE(20, KU&, FIN&)
            NPEDZ& = XVALO(!NROPED)
            Call VERIDESPRE(NPEDZ&, 0, HUBORECU%)
         .MoveNext
         Loop
         Screen.MousePointer = 1
      End If
   End With
   IPEPE.Close
   '
   SQLX$ = "SELECT DISTINCT PEDDETA.NroPed, PEDENCA.Referencia as REFEREN FROM PEDDETA WITH(NOLOCK)"
   SQLX$ = SQLX$ + " INNER JOIN PEDENCA"
   SQLX$ = SQLX$ + " ON PEDDETA.NroPed = PEDENCA.NroPed"
   SQLX$ = SQLX$ + " WHERE PEDDETA.NroPed > 0 "
   SQLX$ = SQLX$ + " AND PEDENCA.Status < 9 "                ' NO ANULADOS
   SQLX$ = SQLX$ + " AND PEDDETA.Status <> 8 "
   SQLX$ = SQLX$ + " AND PEDDETA.Status <> 9 "
   SQLX$ = SQLX$ + " AND PEDDETA.StaDespa < 2 "              ' elimina los cancelados
   SQLX$ = SQLX$ + " AND PEDENCA.CODIEMPR = " & CodiEmp% & " "
   SQLX$ = SQLX$ + " AND PEDDETA.CODIEMPR = " & CodiEmp% & " "
   SQLX$ = SQLX$ + " AND (AcuPrepa IS NULL OR AcuPrepa < CanPed)"
   SQLX$ = SQLX$ + " AND PEDDETA.VenRep = 'V' "             ' Excluye Reparaciones
   SQLX$ = SQLX$ + " ORDER BY PEDDETA.NroPed ASC"
   Set IPEPE = New ADODB.Recordset
   IPEPE.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
   '
   JJ& = 0
   With IPEPE
      On Error Resume Next
      .MoveLast
      If Not Err Then
         On Error GoTo 0
         FIN& = .RecordCount
         .MoveFirst
         Do While Not .EOF
           YY$ = MKS$(!NROPED) + !Referen
           JJ& = JJ& + 1
           Call TRACE(20, JJ&, FIN&)
           Call GRAREG("PEPENPRE", YY$, JJ&)
         .MoveNext
         Loop
      End If
   End With
   '
   Call SETULTREG("PEPENPRE", JJ&)
   Call FRESHECHO("PEPENPRE")
   If JJ& < 1 Then
     Call COMUNI("No hay Pedidos\Pendientes de Preparación.")
     GoTo 99
   End If
   '
3  Screen.MousePointer = 1
   Call SELECHO("PEPENPRE/Indice de Pedidos Pendientes de Preparación")
   NPX& = Val(VALACT1$("PEPENPRE"))
   If NPX& < 1 Then GoTo 99
   '
5  Label7 = TRIM$(Str$(NPX&))
   '
   Call CIERRARCH("*.*")
   Call BLANARCH("!BOLREDET")
   '
   Screen.MousePointer = 11
   VUELTA% = 0
11 NPED& = Val(Label7)
   Call VERIDESPRE(NPED&, 0, HUBORECU%)
   REBLA$ = REGBLAN$("DESPAPED")
   JJ& = 0
   '
   SQLX$ = " SELECT CodiInt, Status, NroClie, CanPed, AcuPrepa, FeSolEnt, NUORITEM FROM PEDDETA WITH(NOLOCK) "
   SQLX$ = SQLX$ + " WHERE NroPed = " & NPED& & " "   ' COINCIDE NÚMERO DE PEDIDO
   SQLX$ = SQLX$ + " order by NuOrItem "
   Dim DESPAPE As ADODB.Recordset
   Set DESPAPE = New ADODB.Recordset
   DESPAPE.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
'   Set DESPAPE = READSET(SQLX$)
   With DESPAPE
      On Error GoTo 0 'Resume Next
      .MoveLast
      If Err Then
         On Error GoTo 0
         Call MENSERR(24, "Imposible Preparar Pedido " + TRIM$(Str$(NPED&)))
         GoTo 3
      End If
      On Error GoTo 0
      .MoveFirst
      Do While Not .EOF
         CI% = !CODIINT
         STATU% = !Status
         NCLIE% = !NROCLIE
         CAPENE# = XVALO(!CanPed)
         CAYAPREP# = XVALO(!AcuPrepa)
         If CAYAPREP# <= CAPENE# - 0.05 Then
           FENTRESO% = CVINT(!FeSolEnt)
           NORITE& = !NuOrItem
           '
           Y$ = REGBLAN$("!BOLREDET")
           Call REPLA(Y$, MKI$(CI%), 83, 2)
           Call REPLA(Y$, Unimed$(CI%), 85, 4)
           Call REPLA(Y$, MKD$(CAPENE#), 89, 8)    ' total a despachar
           Call REPLA(Y$, MKS$(NPED&), 151, 4)
           Call REPLA(Y$, MKI$(FENTRESO%), 155, 2)
           Call REPLA(Y$, MKD$(CAYAPREP#), 163, 8) ' total ya preparado
              SALAPREP# = CAPENE# - CAYAPREP#
              If SALAPREP# < 0 Then SALAPREP# = 0
           Call REPLA(Y$, MKD$(SALAPREP#), 171, 8) ' saldo a preparar
           Call REPLA(Y$, MKS$(NORITE&), 189, 4)
           Call REGAPP("!BOLREDET", Y$)
         End If
      .MoveNext
      Loop
   End With
   '
   Call CIERRARCH("!BOLREDET")
   If ULTREG&("!BOLREDET") < 1 Then
     Call MENSERR(24, "Pedido sin Pendientes de Preparación")
     GoTo 2
   End If
   '
   VTB% = VENTABU%("PREPAPED/NC/TITUPAN=Preparación de Materiales Vendidos - " + TRIM$(rasocli$(NCLIE%)))
   If VTB% < 0 Or ULTREG&("!BOLREDET") < 1 Then
     GoTo 3
   End If
   '
   For YX& = 1 To ULTREG&("!BOLREDET")
     YY$ = REGLEIDO$("!BOLREDET", YX&)
     CI% = CVI(Mid$(YY$, 83, 2))
     CAN# = CVD(Mid$(YY$, 163, 8)) + CVD(Mid$(YY$, 171, 8))
     If CI% > 0 Then
       If CI% <= NUMAS% Then
         If CAN# > 0.005 Then
           GoTo 15
         End If
       End If
     End If
   Next YX&
   Call MENSERR(24, "Despacho Cancelado !!!\Faltan Cantidades o Códigos.")
   GoTo 3
   '
15 CONTA% = ULTREG&("!BOLREDET")
   If CONTA% > 0 Then
     Screen.MousePointer = 11  'Call BLIMESS("Registrando Recepcion ...")
     '
     TTXXX$ = "Materiales Vendidos"
     '
16   Call BLOQARCH("MOVISTO")
     ANOTA$ = ""
     TTXYZ$ = Space$(76)
     HOII% = HOY(HOS$)
     Call REPLA(TTXYZ$, HOS$, 1, 8)
     Call REPLA(TTXYZ$, "Preparación de " + TTXXX$ + ":", 11, 56)
     Call REPLA(TTXYZ$, USERTER$, 69, 24)
     ANOTA$ = ANOTA$ + TTXYZ$ + Chr$(10) + Chr$(13)
     '
     Call COPYSTRU("BOLREDET", "MATPREPA")
     '
     AINDI% = 0: JJ& = 0
     For YX& = 1 To ULTREG&("!BOLREDET")
       YY$ = REGLEIDO$("!BOLREDET", YX&)
       CI% = CVI(Mid$(YY$, 83, 2))
       CAPENET = CVD(Mid$(YY$, 89, 8))
       CAPREP# = CVD(Mid$(YY$, 171, 8))
       CAPRES = CAPREP#
       LOTE$ = Mid$(YY$, 69, 12)
       If CI% > 0 Then
         If CI% <= NUMAS% Then
            NORII% = CVS(Mid$(YY$, 189, 4))
            If CAPREP# > 0.005 Then
               ' ANOTAR EN ANOTAPED
               TTXX$ = Space$(76)
               CIXI% = CI%
               Call REPLA(TTXX$, "- " + CODEXT$(CIXI%), 11, 18)
               Call REPLA(TTXX$, DESCRIT0$(CI%), 29, 38)
               Call REPLA(TTXX$, FORMATNUM$(CAPREP#, "F10.1"), 57, 10)
               ANOTA$ = ANOTA$ + Left$(TTXX$, 96) + Chr$(10) + Chr$(13)
               '
               ' REALIZAR TRANSFERENCIA DE STOCK
               CAVER$ = TRIM$(Str$(NPED&))
               While Len(CAVER$) > 6: CAVER$ = Mid$(CAVER$, 2): Wend
               While Len(CAVER$) < 6: CAVER$ = "0" + CAVER$: Wend
               IK$ = TRIM$(Str$(NORII%))
               CAVERI$ = CAVER$ + "-" + IK$
               FEMOVI% = HOY(HO$)
               DOPRI$ = "VP." & CStr(NPED&)
               DOSEC$ = CAVERI$
               REFMO$ = "Prep.O/V-" + CAVERI$
               CAPREPA = CAPREP#
               Call MOVISTOK(FEMOVI%, CI%, LOTE$, "PP", DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFMO$, 0, VUNI#, 1, NCLIE%, DEPEX%, CAPREPA, NPED&, NORII%)
               Call MOVISTOK(FEMOVI%, CI%, LOTE$, "VP", DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFMO$, 0, VUNI#, 1, NCLIE%, DETER%, (-1) * CAPREPA, NPED&, NORII%)
               '
               Call VERIDESPRE(NPED&, NORII%, HUBORECU%)
               '
               ' GRABA INFORMACION PREPARADA PARA EL LISTADO DE MATERIALES PREPARADOS
               JJ& = JJ& + 1
               Call GRAREG("MATPREPA", YY$, JJ&)
            End If
         End If
       End If
     Next YX&
     Call CIERRARCH("MOVISTO")
     Call SETULTREG("MATPREPA", JJ&)
     '
     ANOTA$ = ANOTA$ + Left$(Space$(10) + String$(56, "-"), 96) + Chr$(10) + Chr$(13)
     ANOTA$ = ANOTA$ + String$(88, "=") + Chr$(10) + Chr$(13)
     '
     strSQL = "Select * from Pedenca where NroPed=" & NPED&
     Dim PedEnca As ADODB.Recordset
     Set PedEnca = New ADODB.Recordset
     PedEnca.Open FILTSQL$(strSQL), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
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
     '
     ' EMITE LISTADO DE MATERIALES PREPARADOS A DECISION DEL USUARIO
     If COMALTER%("¿Desea Imprimir la Lista de Materiales Preparados?\\No, Continuar\Si, Imprimir") = 2 Then
        Call HEADERS("MATPREPA", "")
        Call HEADERS("MATPREPA", "Pedido: " & NPED&)
        Call HEADERS("MATPREPA", "Cliente: " & rasocli$(NCLIE%))
        Call FINAL("*MATPREPA")
     End If
     '
     Screen.MousePointer = 1
     Call ETIDESPA(NPED&)
     '
   End If
   '
   GoTo 3
   '
99 Call CIERRARCH("*.*")
   Screen.MousePointer = 1
   Command46.Enabled = True
   '
End Sub

Private Sub Command61_Click()
   Command61.Enabled = False
Call COMUNI("No disponible por el momento")
GoTo 99
   FIN& = ULTREG&("PEDIPREP")
   '
   JJ& = 0
   Call COPYSTRU("PEDIPREP", "PEDIPREQ")
   For KU& = 1 To FIN&
     XX$ = REGLEIDO$("PEDIPREP", KU&)
     '
     NROPED& = CVS(Mid$(XX$, 17, 4))
     CANPREP = CVS(Mid$(XX$, 91, 4))
     CI% = CVI(Mid$(XX$, 49, 2))
'AA1$ = CODEXT$(CI%)
     '
     CAPEPE = 0
     RFG$ = RECFILT$("DESPAPED", 1, MKS$(NROPED&))
     For KX& = 1 To Len(RFG$) Step 4
       REPEDA& = CVS(Mid$(RFG$, KX&, 4))
       X$ = REGLEIDO$("DESPAPED", REPEDA&)
       If CVI(Mid$(X$, 33, 2)) = CI% Then
         If Asc(Mid$(X$, 27, 1)) < 2 Then
           If Mid$(X$, 51, 1) <> "R" Then ' no es reparacion
             CAPEN = CVS(Mid$(X$, 75, 4))
             If CAPEN < 0 Then CAPEN = 0
             CAPEPE = CAPEPE + CAPEN
           End If
         End If
       End If
     Next KX&
     If CAPEPE < CANPREP Then CANPREP = CAPEPE
     '
     If CANPREP > 0.05 Then
       JJ& = JJ& + 1
       Call REPLA(XX$, MKS$(CANPREP), 91, 4)
       Call GRAREG("PEDIPREQ", XX$, JJ&)
     End If
     '
   Next KU&
   Call SETULTREG("PEDIPREQ", JJ&)
   '
   Call FILESORT("PEDIPREQ", "LPEDPREP", 3, 3, "ASC")
   If Option1.Value = True Then
       Call FINAL("*COPEPREP")
     Else
       Call FINAL("*LIPEPREP")
   End If
99 Command61.Enabled = True
End Sub
'
Private Sub Command62_Click()
    Command62.Enabled = False
    Call CALPEPREP
    Call FILESORT("LIPEDPE", "", 1, 1, "ASC")
    '
    If Option1.Value = True Then
       sali$ = "*COPRECL"
     Else
       sali$ = "*LIPRECL"
    End If
    Call FINAL(sali$)
    '
    Command62.Enabled = True
End Sub

Private Sub Command63_Click()
    Command63.Enabled = False
    Call CALPEPREP
    Call FILESORT("LIPEDPE", "", 13, 13, "ASC")
    '
    If Option1.Value = True Then
       sali$ = "*COPREPE"
     Else
       sali$ = "*LIPREPE"
    End If
    Call FINAL(sali$)
    '
99  Command63.Enabled = True
    '
End Sub

Sub CALPEPREP()
   '
   Screen.MousePointer = 11
   Call HEADERS("LIPEDPE", "")
   '
   If DEPEX% < 1 Then
     DEPEX% = Int(Val(Control$("", "DEPOEXPE")))
     If DEPEX% < 1 Then
       Call MENSERR(24, "Proceso Imposible.\  \Falta Configurar Depósitos\de Expedición\y Productos Terminados.")
       GoTo 99
     End If
   End If
   '
   Screen.MousePointer = 11
   SQLX$ = "SELECT * FROM PEDDETA WITH(NOLOCK)"
   SQLX$ = SQLX$ + " INNER JOIN PEDENCA"
   SQLX$ = SQLX$ + " ON PEDDETA.NroPed = PEDENCA.NroPed"
   SQLX$ = SQLX$ + " WHERE PEDDETA.NroPed > 0 "
   SQLX$ = SQLX$ + " AND PEDENCA.Status < 9 "                ' NO ANULADOS
   SQLX$ = SQLX$ + " AND PEDDETA.Status <> 8 "
   SQLX$ = SQLX$ + " AND PEDDETA.Status <> 9 "
   SQLX$ = SQLX$ + " AND PEDDETA.StaDespa < 2 "              ' elimina los cancelados
   SQLX$ = SQLX$ + " AND PEDENCA.CODIEMPR = " & CodiEmp% & " "
   SQLX$ = SQLX$ + " AND PEDDETA.CODIEMPR = " & CodiEmp% & " "
   SQLX$ = SQLX$ + " AND AcuPrepa > AcuDespa"
   SQLX$ = SQLX$ + " AND PEDDETA.VenRep = 'V' "             ' Excluye Reparaciones
   SQLX$ = SQLX$ + " ORDER BY PEDDETA.NroPed ASC, PEDDETA.NUORITEM ASC "
   'Set IPEPE = New ADODB.Recordset
   'IPEPE.Open FILTSQL$(sqlx$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
   Dim IPEPE As ADODB.Recordset
   Set IPEPE = READSET(SQLX$)
   '
   JJ& = 0
   ULTARTI% = NUMAS%
   REBLA$ = REGBLAN$("LIPEDPE")
   With IPEPE
      On Error Resume Next
      .MoveLast
      If Not Err Then
         On Error GoTo 0
         FIN& = .RecordCount
         .MoveFirst
         Do While Not .EOF
           JJ& = JJ& + 1
           Call TRACE(20, JJ&, FIN&)
           STATX% = !Status
           REPARX% = 0
           NC% = !NROCLIE
           CI% = !CODIINT
           CAPEDID = !CanPed  ' CANTIDAD PEDIDA
           If CI% > 0 And CI% <= ULTARTI% Then
              CAYEN = !AcuDespa
              CAPRE = !AcuPrepa - CAYEN
              If CAPRE < 0 Then CAPRE = 0
              '
              If CAPRE > 0.05 Then
                CASAL = CAPEDID - CAYEN
                If CASAL > 0.05 Then
                  PRUNI = !PreUnid
                  FENTRESOL% = CVINT(!FeSolEnt)
                  Y$ = REBLA$
                  '
                  Call REPLA(Y$, MKI$(NC%), 1, 2)
                  Call REPLA(Y$, MKI$(!CODIINT), 3, 2)
                  Call REPLA(Y$, !TCOLOR, 5, 16)
                  Call REPLA(Y$, !TTALLE, 21, 8)
                  Call REPLA(Y$, MKS$(!TLargo), 29, 4)
                  Call REPLA(Y$, MKS$(!TAncho), 33, 4)
                  Call REPLA(Y$, MKS$(!TEspesor), 37, 4)
                  Call REPLA(Y$, MKS$(!PreList), 41, 4)
                  Call REPLA(Y$, MKI$(NC%), 45, 2)
                  Call REPLA(Y$, MKS$(!NROPED), 47, 4)
                  Call REPLA(Y$, !NROOCOM, 53, 16)
                  Call REPLA(Y$, MKI$(FENTRESOL%), 69, 2)
                  Call REPLA(Y$, Chr$(STATX%), 71, 1)
                  Call REPLA(Y$, Chr$(!ListPre), 72, 1)
                  Call REPLA(Y$, Chr$(!CondPag), 73, 1)
                  Call REPLA(Y$, Chr$(!VENDED), 74, 1)
                  Call REPLA(Y$, Unimed$(CI%), 77, 4)
                  Call REPLA(Y$, MKS$(!CanPed), 81, 4)
                  Call REPLA(Y$, MKS$(CAYEN), 97, 4)
                  Call REPLA(Y$, MKS$(CASAL), 101, 4)
                  Call REPLA(Y$, MKS$(CAPRE), 111, 4)
                  '
                  JK& = JK& + 1
                  Call GRAREG("LIPEDPE", Y$, JK&)
                End If
              End If
           End If
         .MoveNext
         Loop
         '
      End If
   End With
   '
30 Call SETULTREG("LIPEDPE", JK&)
   Call CIERRARCH("LIPEDPE")
   '
99 Screen.MousePointer = 1
   Exit Sub
   '
End Sub
'
Private Sub Command64_Click()
    Command64.Enabled = False
    '
    Call CALPEPREP
    Call FILESORT("LIPEDPE", "", 3, 3, "ASC")
    '
    If Option1.Value = True Then
       sali$ = "*COPREMA"
     Else
       sali$ = "*LIPREMA"
    End If
    Call FINAL(sali$)
    '
    Command64.Enabled = True
End Sub

Private Sub Form_Load()
   '
   UTILIZA_SKIN% = 1
   Call APLICACIONSKINS(Me)
   Height = 5300
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   '
End Sub

Sub ETIDESPA(NPEDI&)
   '
   Static POSI%
   Dim PRX As Object
   '
   On Error Resume Next
   IMPREDETX$ = Printer.DeviceName
   On Error GoTo 0
   '
   POSI% = 1
   Call APERBASDAT("PEDCLI")
   SQLX$ = "SELECT * FROM PEDENCA"
   SQLX$ = SQLX$ + " WHERE NroPed = " & NPEDI& & ""
   Dim PEDDDENCA As New ADODB.Recordset
   Set PEDDDENCA = New ADODB.Recordset
   PEDDDENCA.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
   '
   With PEDDDENCA
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
          On Error GoTo 0
          NCLIE% = !NROCLIE
          If NCLIE% > 0 Then
            VDEF$ = Space$(124) + String$(4, 0) + Chr$(POSI%)
            Call REPLA(VDEF$, rasocli$(NCLIE%), 1, 36)
              CAENTRE$ = SAFETEXT$(!DOMIENT)
            Call REPLA(VDEF$, CAENTRE$, 37, 36)
              LOENTRE$ = SAFETEXT$(!LOCAENT)
            Call REPLA(VDEF$, LOENTRE$, 73, 36)
              NOCOM$ = SAFETEXT$(!NROOCOM)
            Call REPLA(VDEF$, NOCOM$, 109, 16)
            Call REPLA(VDEF$, MKI$(1), 125, 2)
            Call REPLA(VDEF$, MKI$(1), 127, 2)
            '
            KKI% = 0
10          DEP1% = CInt(Val(Control$("", "DEPOEXPE")))
            DEP2% = CInt(Val(Control$("", "DEPOTERM")))
            '
            Call REPLA(VDEF$, Chr$(DEP1%), 130, 1)
            Call REPLA(VDEF$, Chr$(DEP2%), 131, 1)
            Call REPLA(VDEF$, Chr$(POSI%), 129, 1)
            '
            OBX$ = OBJPLA$("ETIQ-DESPA", VDEF$)
            If OBX$ <> "" Then
              FORIPRE$ = TRIM$(Control$("", "ETIDESP"))
              If FORIPRE$ = "" Then FORIPRE$ = "ETIQ-EXP"
              PORTUREC$ = TRIM$(Control$(IDENTER$, "PUERDES"))
              If PORTUREC$ = "" Then PORTUREC$ = Control$("*", "PUERDES")
              On Error Resume Next
              Printer.TrackDefault = False
              On Error GoTo 0
              If PORTUREC$ <> "" Then
                NEWPORT$ = PORTUREC$ + "/NOMESS"
                Call SETPRINPORT(NEWPORT$)
                If NEWPORT$ <> "" Then
                  Call SETPRINPORT("RESTORE")
                End If
                PRINTERPORT$ = PORTUREC$
              End If
              '
              VDEF$ = OBX$
              If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") Then
                HOII% = HOY(HOS$)
                FECHDOC$ = HOS$
                NUMEDOC$ = TRIM$(Mid$(OBX$, 109, 16))
                TIPODOC$ = "Etiqueta de Expedición"
                TOBUL% = CVI(Mid$(OBX$, 127, 2))
                DESTIDOC% = 0
                '
20              CACOLTAB1% = 0
                CODPARAM$(1) = "TRANSPOR": DOCPARAM$(1) = TRIM$(Mid$(OBX$, 1, 36))
                CODPARAM$(2) = "DOMI-TRA": DOCPARAM$(2) = TRIM$(Mid$(OBX$, 37, 36))
                CODPARAM$(3) = "LOCA-TRA": DOCPARAM$(3) = TRIM$(Mid$(OBX$, 73, 36))
                CODPARAM$(4) = "NRO-OCOM": DOCPARAM$(4) = TRIM$(Mid$(OBX$, 109, 16))
                CODPARAM$(5) = "REF-MAT1": DOCPARAM$(5) = TRIM$(Str$(CVI(Mid$(OBX$, 125, 2))))
                CODPARAM$(6) = "REF-MAT2": DOCPARAM$(6) = TRIM$(Str$(CVI(Mid$(OBX$, 127, 2))))
                CODPARAM$(7) = "NRO-AENT": DOCPARAM$(7) = TRIM$(Str$(NPEDI&))
                CAPARDOC% = 7
                '
                CANCELPRINT% = 0
                NOCONFIRM% = 1
                KKI% = KKI% + 1
                CODPARAM$(5) = "REF-MAT1": DOCPARAM$(5) = TRIM$(Str$(KKI%))
                Call PRINTDOC("@" + FORIPRE$) ' Etiqueta de Expedición
                If KKI% < TOBUL% Then GoTo 20
                '
              End If
              Call SETPRINPORT("RESTORE")
            End If
          End If
      End If
   End With
   '
   PRINTERPORT$ = ""
   '
End Sub

Sub RECUDESPA(NPEDRECU&)
   Screen.MousePointer = 11
   Call FILTERDELETE("DESPAPED", 1, MKS$(NPEDRECU&))
   RFGH$ = RECFILT$("PEDDETA", 1, MKS$(NPEDRECU&))
   For KKUI& = 1 To Len(RFGH$) Step 4
     REPEDE& = CVS(Mid$(RFGH$, KKUI&, 4))
     PDC$ = REGLEIDO$("PEDDETA", REPEDE&)
     If Asc(Mid$(PDC$, 27, 1)) <> 9 Then
         CAN = CVS(Mid$(PDC$, 79, 4))
       Call REPLA(PDC$, Chr$(1), 27, 1)
       Call REPLA(PDC$, String$(38, 0), 71, 38)
       Call REPLA(PDC$, MKS$(CAN), 75, 4)
       Call REPLA(PDC$, MKS$(CAN), 99, 4)
       Call REGAPP("DESPAPED", PDC$)
     End If
   Next KKUI&
   Call CIERRARCH("DESPAPED")
   Screen.MousePointer = 1
End Sub


