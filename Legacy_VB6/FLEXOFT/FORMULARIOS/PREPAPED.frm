VERSION 5.00
Begin VB.Form PREPAPED 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Preparación de Pedidos de Clientes"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5880
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
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00404080&
      Height          =   7260
      Left            =   4935
      ScaleHeight     =   7200
      ScaleWidth      =   1425
      TabIndex        =   27
      Top             =   0
      Width           =   1485
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
         Left            =   150
         Picture         =   "PREPAPED.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Control Depósito de Expedición"
         Top             =   2520
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
         Picture         =   "PREPAPED.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   34
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
         Picture         =   "PREPAPED.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   31
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
         Picture         =   "PREPAPED.frx":0896
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   945
         Width           =   600
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   150
         ScaleHeight     =   75
         ScaleWidth      =   540
         TabIndex        =   28
         Top             =   3885
         Width           =   600
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   29
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -280
         Picture         =   "PREPAPED.frx":0BA0
         Top             =   4320
         Width           =   1515
      End
      Begin VB.Line Line1 
         X1              =   -105
         X2              =   1195
         Y1              =   4725
         Y2              =   4725
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Consultas y Listados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2190
      Left            =   150
      TabIndex        =   8
      Top             =   2600
      Width           =   4635
      Begin VB.PictureBox Picture6 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   420
         ScaleHeight     =   585
         ScaleWidth      =   3735
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
            Height          =   600
            Left            =   3150
            Picture         =   "PREPAPED.frx":2AC2
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
            Height          =   600
            Left            =   2550
            Picture         =   "PREPAPED.frx":2DCC
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
            Height          =   600
            Left            =   0
            Picture         =   "PREPAPED.frx":320E
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
            Height          =   600
            Left            =   1950
            Picture         =   "PREPAPED.frx":3518
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
               Name            =   "MS Sans Serif"
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
         BackColor       =   &H00C0E0FF&
         Caption         =   "Dispositivo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
            BackColor       =   &H00C0E0FF&
            Caption         =   "Pantalla"
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
            Left            =   525
            TabIndex        =   11
            Top             =   270
            Value           =   -1  'True
            Width           =   1065
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Impresión"
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
            Left            =   2205
            TabIndex        =   10
            Top             =   270
            Width           =   1170
         End
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Transacciones "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2265
      Left            =   150
      TabIndex        =   0
      Top             =   210
      Width           =   4635
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Index           =   3
         Left            =   1465
         ScaleHeight     =   585
         ScaleWidth      =   2370
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
            Height          =   600
            Left            =   0
            Picture         =   "PREPAPED.frx":3822
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
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
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
         BackColor       =   &H80000016&
         Height          =   645
         Index           =   2
         Left            =   1100
         ScaleHeight     =   585
         ScaleWidth      =   2370
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
            Height          =   600
            Left            =   0
            Picture         =   "PREPAPED.frx":396C
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
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
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
         BackColor       =   &H80000016&
         Height          =   645
         Index           =   0
         Left            =   735
         ScaleHeight     =   585
         ScaleWidth      =   2370
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
            Height          =   600
            Left            =   0
            Picture         =   "PREPAPED.frx":3C76
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
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
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
      Height          =   1905
      Left            =   150
      TabIndex        =   32
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
         TabIndex        =   33
         Top             =   315
         Width           =   1065
      End
   End
End
Attribute VB_Name = "PREPAPED"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   If DERACCE%("PREPAPED", "Preparar Pedidos de Clientes") < 2 Then
     GoTo 99
   End If
   '
   Dim CAPRENET(32767), CASTOEXP(32767)
   DEPEX% = Int(Val(CONTROL$("", "DEPOEXPE")))
   DETER% = Int(Val(CONTROL$("", "DEPOTERM")))
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
     Call VERIDESPRE(REDE&, HUBORECU%)
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
     TTXYZ$ = Space$(76)
     HOII% = HOY(HOS$)
     Call REPLA(TTXYZ$, HOS$, 1, 8)
     Call REPLA(TTXYZ$, "Cancelacion General Preparados y Despacho ", 11, 56)
     Call REPLA(TTXYZ$, USERTER$, 69, 24)
     Call REGAPP("ANOTAPED", MKS$(NPXX&) + TTXYZ$)
     TTXYZ$ = Space$(76)
     Call REPLA(TTXYZ$, "por Aparente Sobre-Embarque. ", 11, 56)
     Call REGAPP("ANOTAPED", MKS$(NPXX&) + TTXYZ$)
     Call REGAPP("ANOTAPED", MKS$(NPXX&) + Space$(88))
     Call REGAPP("ANOTAPED", MKS$(NPXX&) + String$(88, "="))
     Call CIERRARCH("ANOTAPED")
     Call FRESHECHO("ANOTAPED")
     Call COMUNI("Vuelva a Ingresar\a Preparacion de Pedidos")
     Call FINAL("")
     '
   End If
Return
'
End Sub

'
Private Sub Command26_Click()
   Command26.Enabled = False
   CFGDESPA.Show 1
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
    RFF$ = RECFILT$("MOVISTO", 5, "PP") ' SELECCIONA MOVIMIENTOS DE PREPARACION DE PEDIDOS
    RFF$ = RFF$ + RECFILT$("MOVISTO", 5, "DM")
    JJ& = 0: FIN& = Len(RFF$)
    For U& = 1 To FIN& Step 4
      Call TRACE(20, U&, FIN&)
      REMO& = CVS(Mid$(RFF$, U&, 4))
      MM$ = REGLEIDO$("MOVISTO", REMO&)
      If Asc(Mid$(MM$, 83, 1)) = DEPX% Then      ' MOVIMIENTO EN DEPODITO DE EXPEDICION
        IDPEDIT$ = Mid$(MM$, 112, 6)             ' LEE PEDIDO E ITEM
        If CVS(Left$(IDPEDIT$, 4)) >= 0.5 Then
          For KU& = 1 To JJ&
            If ITPEDI$(KU&) = IDPEDIT$ Then
              CAPENDE(KU&) = CAPENDE(KU&) + CVD(Mid$(MM$, 88, 8)) - CVD(Mid$(MM$, 96, 8))
              GoTo 19
            End If
          Next KU&
          JJ& = JJ& + 1
          If JJ& > MAXPED& Then
            MAXPED& = MAXPED& + 4096
            ReDim Preserve ITPEDI$(MAXPED&), CAPENDE(MAXPED&), NUPEDI&(MAXPED&)
          End If
          ITPEDI$(JJ&) = IDPEDIT$
          CAPENDE(JJ&) = CAPENDE(JJ&) + CVD(Mid$(MM$, 88, 8)) - CVD(Mid$(MM$, 96, 8))
          NUPEDI&(JJ&) = CVS(Left$(IDPEDIT$, 4))
        End If
      End If
19  Next U&
    '
    FIN& = ULTREG&("DESPAPED")
    NPXA& = 0: J2& = 0
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("DESPAPED", U&)
      NPX& = CVS(Left$(XX$, 4))
      If NPX& <> NPXA& Then
        For J1& = 1 To JJ&
          If NUPEDI&(J1&) = NPX& Then
            If CAPENDE(J1&) >= 0.05 Then
              NPXA& = NPX&
              NCLIE% = CVI(Mid$(XX$, 7, 2))
              NUOCO$ = TRIM$(Mid$(XX$, 9, 16))
              FEPE = CVI(Mid$(XX$, 5, 2))
              REPAR$ = "": If Mid$(XX$, 51, 1) = "R" Then REPAR$ = " (Reparación)"
              YY$ = MKS$(NPX&) + TRIM$(RASOCLI$(NCLIE%))
              If NUOCO$ <> "" Then YY$ = YY$ + " - " + NUOCO$
              YY$ = YY$ + " - " + FECHATEX$(FEPE) + REPAR$
              J2& = J2& + 1
              Call GRAREG("IPEDPREP", YY$, J2&)
              GoTo 29
            End If
          End If
        Next J1&
      End If
29  Next U&
    Screen.MousePointer = 1
    '
    Call SETULTREG("IPEDPREP", J2&)
    Call FRESHECHO("IPEDPREP")
    If J2& < 1 Then
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
     DEPEX% = Int(Val(CONTROL$("", "DEPOEXPE")))
     DETER% = Int(Val(CONTROL$("", "DEPOTERM")))
     If DEPEX% < 1 Or DETER% < 1 Or DEPEX% = DETER% Then
       Call MENSERR(24, "Proceso Imposible.\  \Falta Configurar Depósitos\de Expedición\y Productos Terminados.")
       GoTo 99
     End If
   End If
   '
   Screen.MousePointer = 11
   Call COPYSTRU("INDIPEDI", "IPEDPREP")
   RFF$ = RECFILT$("DESPAPED", 7, Chr$(1))
   FIN& = Len(RFF$)
   '
   Screen.MousePointer = 11
   For U& = 1 To FIN& Step 4
     Call TRACE(20, U&, FIN&)
     REDE& = CVS(Mid$(RFF$, U&, 4))
     ' AQUI VERIFICAR ESTADO DE PREPARACION DEL REGISTRO Y PASARLO A STATUS 3 SI ESTA CUMPLIDO
     Call VERIDESPRE(REDE&, HUBORECU%)
   Next U&
   '
   RFF$ = RECFILT$("DESPAPED", 7, Chr$(11))
   NPXA& = 0: JJ& = 0: FIN& = Len(RFF$)
   For U& = 1 To FIN& Step 4
     REDE& = CVS(Mid$(RFF$, U&, 4))
     XX$ = REGLEIDO$("DESPAPED", REDE&)
     NPX& = CVS(Left$(XX$, 4))
     CAPEDIT = CVS(Mid$(XX$, 75, 4))
     ACUPRE = CVS(Mid$(XX$, 79, 4))
     ACUDES = CVS(Mid$(XX$, 95, 4))
     SALPENDES = ACUPRE - ACUDES
     If CAPEDIT < ACUPRE Then
       SALPENDES = CAPEDIT - ACUDES
     End If
     '
     If SALPENDES >= 0.05 Then
       If NPX& <> NPXA& Then
         NPXA& = NPX&
         NCLIE% = CVI(Mid$(XX$, 7, 2))
         NUOCO$ = TRIM$(Mid$(XX$, 9, 16))
         FEPE = CVI(Mid$(XX$, 5, 2))
         'REPAR$ = "": If Mid$(XX$, 51, 1) = "R" Then REPAR$ = " (Reparación)"
         If Mid$(XX$, 51, 1) <> "R" Then
           YY$ = MKS$(NPX&) + TRIM$(RASOCLI$(NCLIE%))
           If NUOCO$ <> "" Then YY$ = YY$ + " - " + NUOCO$
           YY$ = YY$ + " - " + FECHATEX$(FEPE) + REPAR$
           JJ& = JJ& + 1
           Call GRAREG("IPEDPREP", YY$, JJ&)
         End If
       End If
     End If
29 Next U&
   '
   Call SETULTREG("IPEDPREP", JJ&)
   Call FILESORT("IPEDPREP", "", 1, 1)
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
   DEPORIG% = DEPOCON%
   DEPOSI$ = ECOARCH$("TADEPOSI", Chr$(DEPORIG%))
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
   RFF$ = RECFILT$("DESPAPED", 1, MKS$(NPX&))
   '
   If Len(RFF$) < 4 Then
     Call MENSERR(24, "Imposible Des-Preparar Pedido " + TRIM$(Str$(NPED&)))
     GoTo 3
   End If
   '
   ' LO QUE SIGUE ES PARA PEDIDOS DE VENTA
   For U& = 1 To Len(RFF$) Step 4
      REDE& = CVS(Mid$(RFF$, U&, 4))
      XX$ = REGLEIDO$("DESPAPED", REDE&)
      STATU% = Asc(Mid$(XX$, 27, 1))
      'If STATU% < 2 Then
      CI% = CVI(Mid$(XX$, 33, 2))
      NCLIE% = CVI(Mid$(XX$, 7, 2))
        ACUPRE = CVS(Mid$(XX$, 79, 4))
        ACUDES = CVS(Mid$(XX$, 95, 4))
      CAPENE# = ACUPRE - ACUDES
      FENTRESO% = CVI(Mid$(XX$, 25, 2))
      '
      If CAPENE# >= 0.05 Then
        Y$ = REGBLAN$("!BOLREDET")
        Call REPLA(Y$, MKI$(CI%), 83, 2)
        Call REPLA(Y$, UNIMED$(CI%), 85, 4)
        Call REPLA(Y$, MKD$(CAPENE#), 89, 8)
        Call REPLA(Y$, MKS$(NPX&), 151, 4)
        Call REPLA(Y$, MKI$(FENTRESO%), 155, 2)
        Call REPLA(Y$, MKD$(0), 163, 8)
        Call REPLA(Y$, MKS$(REDE&), 189, 4)
        Call REGAPP("!BOLREDET", Y$)
      End If
    Next U&
    '
12  VTB% = VENTABU%("DESPAPED/NC/TITUPAN=Cancelación de Preparación - " + TRIM$(RASOCLI$(NCLIE%)))
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
      Call BLOQARCH("MOVISTO")
      Call BLOQARCH("ANOTAPED")
      Screen.MousePointer = 11  'Call BLIMESS("Registrando Recepcion ...")
      '
      TTXXX$ = "Materiales Vendidos"
      '
16    TTXYZ$ = Space$(76)
      HOII% = HOY(HOS$)
      Call REPLA(TTXYZ$, HOS$, 1, 8)
      Call REPLA(TTXYZ$, "Cancela Preparación de " + TTXXX$ + ":", 11, 56)
      Call REPLA(TTXYZ$, USERTER$, 69, 24)
      Call REGAPP("ANOTAPED", MKS$(NPX&) + TTXYZ$)
      Call REGAPP("ANOTAPED", MKS$(NPX&) + Space$(88))
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
              REDE& = CVS(Mid$(YY$, 189, 4))
              ZY$ = REGLEIDO$("DESPAPED", REDE&)
              ' ANOTAR EN DESPAPED
              TOTAPREP = CVS(Mid$(ZY$, 79, 4)) - CAPREP#
              If TOTAPREP < 0 Then TOTAPREP = 0
              Call REPLA(ZY$, MKS$(TOTAPREP), 79, 4)
              Call REPLA(ZY$, Chr$(1), 27, 1)
              Call GRAREG("DESPAPED", ZY$, REDE&)
              CAVERI$ = Mid$(ZY$, 113, 10)
              '
              ' ANOTAR EN ANOTAPED
              TTXX$ = Space$(76)
              CIXI% = CI%
              Call REPLA(TTXX$, "- " + CODEXT$(CIXI%), 11, 18)
              Call REPLA(TTXX$, DESCRIT0$(CIXI%), 29, 38)
              Call REPLA(TTXX$, CSTRING$(MKS$(CAPRES), 3, 10, 1, 2), 57, 10)
              Call REGAPP("ANOTAPED", Left$(MKS$(NPX&) + TTXX$, 96))
              ' REALIZAR TRANSFERENCIA DE STOCK
              '
              FEMOVI% = HOY(HO$)
              DOPRI$ = "VP." & CStr(NPX&)
              DOSEC$ = CAVERI$
              REFMO$ = "Canc.Prep.O/V-" + CAVERI$
              CAPREPA = CAPREP#
              Call MOVISTO(FEMOVI%, CI%, LOTE$, "PP", DOPRI$, DOSEC$, REFMO$, 0, "$", 0, DEPEX%, (-1) * CAPREPA)
              Call MOVISTO(FEMOVI%, CI%, LOTE$, "VP", DOPRI$, DOSEC$, REFMO$, 0, "$", 0, DETER%, CAPREPA)
            End If
          End If
        End If
      Next YX&
      Call CIERRARCH("MOVISTO")
      '
      Call REGAPP("ANOTAPED", Left$(MKS$(NPX&) + Space$(10) + String$(56, "-"), 96))
      Call REGAPP("ANOTAPED", MKS$(NPX&) + String$(88, "="))
      '
      Call CIERRARCH("ANOTAPED")
      Call CIERRARCH("DESPAPED")
      '
      Call FILTERDELETE("IPEDPREP", 1, MKS$(NPX&))
      Call FRESHECHO("IPEDPREP")
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
   DEPEX% = Int(Val(CONTROL$("", "DEPOEXPE")))
   DETER% = Int(Val(CONTROL$("", "DEPOTERM")))
   If DEPEX% < 1 Or DETER% < 1 Or DEPEX% = DETER% Then
     Call MENSERR(24, "Proceso Imposible.\  \Falta Configurar Depósitos\de Expedición\y Productos Terminados.")
     GoTo 99
   End If
   '
   Screen.MousePointer = 11
   Call COPYSTRU("INDIPEDI", "PEPENPRE")
   RFF$ = RECFILT$("DESPAPED", 7, Chr$(1))
   FIN& = Len(RFF$)
   '
   Screen.MousePointer = 11
   For U& = 1 To FIN& Step 4
     Call TRACE(20, U&, FIN&)
     REDE& = CVS(Mid$(RFF$, U&, 4))
     ' AQUI VERIFICAR ESTADO DE PREPARACION DEL REGISTRO Y PASARLO A STATUS 3 SI ESTA CUMPLIDO
     Call VERIDESPRE(REDE&, HUBORECU%)
     'If STADESPE%(REDE&) > 1 Then GoTo 29
   Next U&
   '
   RFF$ = RECFILT$("DESPAPED", 7, Chr$(11))
   NPXA& = 0: JJ& = 0: FIN& = Len(RFF$)
   For U& = 1 To FIN& Step 4
     REDE& = CVS(Mid$(RFF$, U&, 4))
     XX$ = REGLEIDO$("DESPAPED", REDE&)
     NPX& = CVS(Left$(XX$, 4))
'If NPX& = 13448 Then
'AA1 = Asc(Mid$(XX$, 27, 1))
'AA2 = CVS(Mid$(XX$, 75, 4))
'AA3 = CVS(Mid$(XX$, 79, 4))
'AA4 = CVS(Mid$(XX$, 95, 4))
'AAA = BBB
'End If
     
     CAPEDIT = CVS(Mid$(XX$, 75, 4))
     ACUPRE = CVS(Mid$(XX$, 79, 4))
     ACUDES = CVS(Mid$(XX$, 95, 4))
     SALPENPRE = CAPEDIT - ACUPRE
     If SALPENPRE >= 0.05 Then
       If NPX& <> NPXA& Then
         NPXA& = NPX&
         NCLIE% = CVI(Mid$(XX$, 7, 2))
         NUOCO$ = TRIM$(Mid$(XX$, 9, 16))
         FEPE = CVI(Mid$(XX$, 5, 2))
         'REPAR$ = "": If Mid$(XX$, 51, 1) = "R" Then REPAR$ = " (Reparación)"
         If Mid$(XX$, 51, 1) <> "R" Then
           YY$ = MKS$(NPX&) + TRIM$(RASOCLI$(NCLIE%))
           If NUOCO$ <> "" Then YY$ = YY$ + " - " + NUOCO$
           YY$ = YY$ + " - " + FECHATEX$(FEPE) + REPAR$
           JJ& = JJ& + 1
           Call GRAREG("PEPENPRE", YY$, JJ&)
         End If
       End If
     End If
29 Next U&
   '
   Call SETULTREG("PEPENPRE", JJ&)
   Call FILESORT("PEPENPRE", "", 1, 1)
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
   VECO$ = VALACT2$("PEPENPRE")
   '
5  Label7 = TRIM$(Str$(NPX&))
   '
   Call CIERRARCH("*.*")
   Call BLANARCH("!BOLREDET")
   '
   Screen.MousePointer = 11
   VUELTA% = 0
11 NPED& = Val(Label7)
   REBLA$ = REGBLAN$("DESPAPED")
   RFF$ = RECFILT$("DESPAPED", 1, MKS$(NPED&))
   JJ& = 0
   '
   If Len(RFF$) < 4 Then
     Call MENSERR(24, "Imposible Preparar Pedido " + TRIM$(Str$(NPED&)))
     GoTo 3
   End If
   '
   REDE& = CVS(Left(RFF$, 4))
   XX$ = REGLEIDO$("DESPAPED", REDE&)
   CI% = CVI(Mid$(XX$, 33, 2))
   REPARX% = 0
   If Mid$(XX$, 51, 1) = "R" Then   ' ES UNA REPARACIóN
     REPARX% = 1
     Call COMUNI("No Corresponde Preparar\Pedido de Reparación.")
     GoTo 3
   End If
   '
   ' se elimina lo que sigue por pedido de reparación
'     If COMALTER%("Confirma Preparación del Pedido\de Reparación Numero " + TRIM$(Str$(NPED&)) + "\\Si, Preparar\No, Cancelar") = 1 Then
'       For U& = 1 To Len(RFF$) Step 4
'         REDE& = CVS(Mid$(RFF$, U&, 4))
'         XX$ = REGLEIDO$("DESPAPED", REDE&)
'         Call REPLA(XX$, Chr$(5), 27, 1)
'         Call GRAREG("DESPAPED", XX$, REDE&)
'       Next U&
'       Call CIERRARCH("DESPAPED")
'       Call FILTERDELETE("PEPENPRE", 1, MKS$(NPED&))
'       Call FRESHECHO("PEPENPRE")
'       GoSub ANOTAREP
'       Call ETIDESPA(NPED&)
'       Label7 = ""
'       Call COMUNI("Pedido Numero " + TRIM$(Str$(NPED&)) + " Preparado.")
'     End If
'     GoTo 3
   '
   ' LO QUE SIGUE ES PARA PEDIDOS DE VENTA
   '
   Call VERIDESPRE(0, HUBORECU%)
   HUBOREC% = 0
   For U& = 1 To Len(RFF$) Step 4
     REDES& = CVS(Mid$(RFF$, U&, 4))
     Call VERIDESPRE(REDES&, HUBORECU%)
     If HUBORECU% > 0 Then HUBOREC% = 1
   Next U&
   Call CIERRARCH("DESPAPED")
   Screen.MousePointer = 1
   If HUBOREC% > 0 Then
     Call COMUNI("Se Han Recuperado Acumulados\de Preparación y Despacho.")
   End If
   '
   HUBODESPANT% = 0
   VECPEDIT$ = ""
   For U& = 1 To Len(RFF$) Step 4
     REDE& = CVS(Mid$(RFF$, U&, 4))
     XX$ = REGLEIDO$("DESPAPED", REDE&)
     '
     STATU% = Asc(Mid$(XX$, 27, 1))
     CI% = CVI(Mid$(XX$, 33, 2))
     NCLIE% = CVI(Mid$(XX$, 7, 2))
     '
     CAPENE# = CVS(Mid$(XX$, 75, 4))
     CAYAPREP# = CVS(Mid$(XX$, 79, 4))
     If CAYAPREP# <= CAPENE# - 0.05 Then
       FENTRESO% = CVI(Mid$(XX$, 25, 2))
       CAVERI$ = Mid$(XX$, 113, 10)
       '
       Y$ = REGBLAN$("!BOLREDET")
       Call REPLA(Y$, MKI$(CI%), 83, 2)
       Call REPLA(Y$, UNIMED$(CI%), 85, 4)
       Call REPLA(Y$, MKD$(CAPENE#), 89, 8)    ' total a despachar
       Call REPLA(Y$, MKS$(NPED&), 151, 4)
       Call REPLA(Y$, MKI$(FENTRESO%), 155, 2)
       Call REPLA(Y$, MKD$(CAYAPREP#), 163, 8) ' total ya preparado
         SALAPREP# = CAPENE# - CAYAPREP#
         If SALAPREP# < 0 Then SALAPREP# = 0
       Call REPLA(Y$, MKD$(SALAPREP#), 171, 8) ' saldo a preparar
       Call REPLA(Y$, MKS$(REDE&), 189, 4)
       Call REGAPP("!BOLREDET", Y$)
     End If
   Next U&
   Call CIERRARCH("!BOLREDET")
   If ULTREG&("!BOLREDET") < 1 Then
     Call MENSERR(24, "Pedido sin Pendientes de Preparación")
     GoTo 3
   End If
   '
   VTB% = VENTABU%("PREPAPED/NC/TITUPAN=Preparación de Materiales Vendidos - " + TRIM$(RASOCLI$(NCLIE%)))
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
     Call CIERRARCH("MOVISTO")
     Screen.MousePointer = 11  'Call BLIMESS("Registrando Recepcion ...")
     '
     TTXXX$ = "Materiales Vendidos"
     If REPARX% = 1 Then
        TTXXX$ = "Reparación"
        NREP& = 0
        For KU& = ULTREG&("PRESREPA") To 1 Step -1
          XYX$ = REGLEIDO$("PRESREPA", KU&)
          If CVS(Mid$(XYX$, 125, 4)) = NPED& Then
            NREP& = CVS(Left$(XYX$, 4))
            GoTo 16
          End If
        Next KU&
     End If
     '
16   Call BLOQARCH("ANOTAPED")
     Call BLOQARCH("ANOTAREP")
     Call BLOQARCH("MOVISTO")
     TTXYZ$ = Space$(76)
     HOII% = HOY(HOS$)
     Call REPLA(TTXYZ$, HOS$, 1, 8)
     Call REPLA(TTXYZ$, "Preparación de " + TTXXX$ + ":", 11, 56)
     Call REPLA(TTXYZ$, USERTER$, 69, 24)
     Call REGAPP("ANOTAPED", MKS$(NPED&) + TTXYZ$)
     Call REGAPP("ANOTAPED", MKS$(NPED&) + Space$(88))
     If REPARX% = 1 Then
       Call REGAPP("ANOTAREP", MKS$(NREP&) + TTXYZ$)
       Call REGAPP("ANOTAREP", MKS$(NREP&) + Space$(88))
     End If
     '
     AINDI% = 0
     For YX& = 1 To ULTREG&("!BOLREDET")
       YY$ = REGLEIDO$("!BOLREDET", YX&)
       CI% = CVI(Mid$(YY$, 83, 2))
       CAPENET = CVD(Mid$(YY$, 89, 8))
       CAPREP# = CVD(Mid$(YY$, 171, 8))
       CAPRES = CAPREP#
       CANTOTPREP = CVD(Mid$(YY$, 163, 8)) + CAPREP#
       LOTE$ = Mid$(YY$, 69, 12)
       If CI% > 0 Then
         If CI% <= NUMAS% Then
           If CANTOTPREP > 0.005 Then
             REDE& = CVS(Mid$(YY$, 189, 4))
             ZY$ = REGLEIDO$("DESPAPED", REDE&)
             If CAPREP# > 0.005 Then
               ' ANOTAR EN DESPAPED
               TOTAPREP = CVS(Mid$(ZY$, 79, 4)) + CAPREP#
               Call REPLA(ZY$, MKS$(TOTAPREP), 79, 4)
               Call REPLA(ZY$, Chr$(1), 27, 1)
               Call GRAREG("DESPAPED", ZY$, REDE&)
               CAVERI$ = Mid$(ZY$, 113, 10)
               '
               ' ANOTAR EN ANOTAPED
               TTXX$ = Space$(76)
               CIXI% = CI%
               Call REPLA(TTXX$, "- " + CODEXT$(CIXI%), 11, 18)
               Call REPLA(TTXX$, DESCRIT0$(CIXI%), 29, 38)
               Call REPLA(TTXX$, CSTRING$(MKS$(CAPRES), 3, 10, 1, 2), 57, 10)
               Call REGAPP("ANOTAPED", Left$(MKS$(NPED&) + TTXX$, 96))
               If REPARX% = 1 Then
                 Call REGAPP("ANOTAREP", MKS$(NREP&) + TTXX$)
               End If
               '
               ' REALIZAR TRANSFERENCIA DE STOCK
               If REPARX% < 1 Then    ' EXCLUYE REPARACIONES
                 FEMOVI% = HOY(HO$)
                 DOPRI$ = "VP." & CStr(NPED&)
                 DOSEC$ = CAVERI$
                 REFMO$ = "Prep.O/V-" + CAVERI$
                 CAPREPA = CAPREP#
                 Call MOVISTO(FEMOVI%, CI%, LOTE$, "PP", DOPRI$, DOSEC$, REFMO$, 0, "$", 0, DEPEX%, CAPREPA)
                 Call MOVISTO(FEMOVI%, CI%, LOTE$, "VP", DOPRI$, DOSEC$, REFMO$, 0, "$", 0, DETER%, (-1) * CAPREPA)
               End If
             End If
             '
           End If
         End If
       End If
     Next YX&
     Call CIERRARCH("MOVISTO")
     '
     Call REGAPP("ANOTAPED", Left$(MKS$(NPED&) + Space$(10) + String$(56, "-"), 96))
     Call REGAPP("ANOTAPED", MKS$(NPED&) + String$(88, "="))
     If REPARX% = 1 Then
       Call REGAPP("ANOTAREP", Left$(MKS$(NREP&) + Space$(10) + String$(56, "-"), 96))
       Call REGAPP("ANOTAREP", MKS$(NREP&) + String$(88, "="))
     End If
     '
     Call CIERRARCH("ANOTAPED")
     Call CIERRARCH("ANOTAREP")
     Call CIERRARCH("DESPAPED")
     '
     Label7 = ""
     '
     Screen.MousePointer = 1
     Call ETIDESPA(NPED&)
     '
   End If
   '
   Call CIERRARCH("MOVISTO")
   GoTo 3
   '
99 Call CIERRARCH("MOVISTO")
   Call CIERRARCH("*.*")
   Screen.MousePointer = 1
   Command46.Enabled = True
Exit Sub
'
'

ANOTAREP:
   TTXXX$ = "Reparación"
   NREP& = 0
   For KU& = ULTREG&("PRESREPA") To 1 Step -1
     XYX$ = REGLEIDO$("PRESREPA", KU&)
     If CVS(Mid$(XYX$, 125, 4)) = NPED& Then
       NREP& = CVS(Left$(XYX$, 4))
       GoTo 86
     End If
   Next KU&
   '
86 Call BLOQARCH("ANOTAPED")
   Call BLOQARCH("ANOTAREP")
   '
   TTXYZ$ = Space$(76)
   HOII% = HOY(HOS$)
   Call REPLA(TTXYZ$, HOS$, 1, 8)
   Call REPLA(TTXYZ$, "Preparación de " + TTXXX$ + ":", 11, 56)
   Call REPLA(TTXYZ$, USERTER$, 69, 24)
   Call REGAPP("ANOTAPED", MKS$(NPED&) + TTXYZ$)
   Call REGAPP("ANOTAPED", MKS$(NPED&) + Space$(88))
   Call REGAPP("ANOTAREP", MKS$(NREP&) + TTXYZ$)
   Call REGAPP("ANOTAREP", MKS$(NREP&) + Space$(88))
   '
   TTXX$ = Space$(76)
   CIXI% = CI%
   Call REPLA(TTXX$, "- " + CODEXT$(CIXI%), 11, 18)
   Call REPLA(TTXX$, DESCRIT0$(CIXI%), 29, 38)
   Call REPLA(TTXX$, CSTRING$(MKS$(CAPRES), 3, 10, 1, 2), 57, 10)
   Call REGAPP("ANOTAPED", Left$(MKS$(NPED&) + TTXX$, 96))
   Call REGAPP("ANOTAREP", MKS$(NREP&) + TTXX$)
   '
   Call REGAPP("ANOTAPED", Left$(MKS$(NPED&) + Space$(10) + String$(56, "-"), 96))
   Call REGAPP("ANOTAPED", MKS$(NPED&) + String$(88, "="))
   Call REGAPP("ANOTAREP", Left$(MKS$(NREP&) + Space$(10) + String$(56, "-"), 96))
   Call REGAPP("ANOTAREP", MKS$(NREP&) + String$(88, "="))
   '
   Call CIERRARCH("ANOTAPED")
   Call CIERRARCH("ANOTAREP")
Return

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
       Sali$ = "*COPRECL"
     Else
       Sali$ = "*LIPRECL"
    End If
    Call FINAL(Sali$)
    '
    Command62.Enabled = True
End Sub

Private Sub Command63_Click()
    Command63.Enabled = False
    Call CALPEPREP
    Call FILESORT("LIPEDPE", "", 13, 13, "ASC")
    '
    If Option1.Value = True Then
       Sali$ = "*COPREPE"
     Else
       Sali$ = "*LIPREPE"
    End If
    Call FINAL(Sali$)
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
     DEPEX% = Int(Val(CONTROL$("", "DEPOEXPE")))
     If DEPEX% < 1 Then
       Call MENSERR(24, "Proceso Imposible.\  \Falta Configurar Depósitos\de Expedición\y Productos Terminados.")
       GoTo 99
     End If
   End If
   '
   Screen.MousePointer = 11
   Call COPYSTRU("INDIPEDI", "PEPENDES")
   RFF$ = RECFILT$("DESPAPED", 7, Chr$(1))
   FIN& = Len(RFF$)
   '
   Screen.MousePointer = 11
   For U& = 1 To FIN& Step 4
     Call TRACE(20, U&, FIN&)
     REDE& = CVS(Mid$(RFF$, U&, 4))
     ' AQUI VERIFICAR ESTADO DE PREPARACION DEL REGISTRO Y PASARLO A STATUS 3 SI ESTA CUMPLIDO
     Call VERIDESPRE(REDE&, HUBORECU%)
   Next U&
   '
   RFF$ = RECFILT$("DESPAPED", 7, Chr$(11))
   FIN& = Len(RFF$)
   JK& = 0
   For KU& = 1 To FIN& Step 4
     Call TRACE(20, KU&, FIN&)
     U& = CVS(Mid$(RFF$, KU&, 4))
     X$ = REGLEIDO$("DESPAPED", U&)
     STATX% = Asc(Mid$(X$, 27, 1))
     REPARX% = 0: If Mid$(X$, 51, 1) = "R" Then REPARX% = 1
     '
     If REPARX% < 1 Then
       NC% = CVI(Mid$(X$, 7, 2))
       CI% = CVI(Mid$(X$, 33, 2))
       CAPEDID = CVS(Mid$(X$, 75, 4))  ' CANTIDAD PEDIDA
       If CI% > 0 And CI% <= NUMAS% Then
         CAYEN = CVS(Mid$(X$, 95, 4))
         CAPRE = CVS(Mid$(X$, 79, 4)) - CAYEN
         If CAPRE < 0 Then CAPRE = 0
         '
         If CAPRE > 0.05 Then
           CASAL = CAPEDID - CAYEN
           If CASAL > 0.05 Then
             PRUNI = CVS(Mid$(X$, 83, 4))
             X1$ = X$
             Y$ = String$(46, 0) + Left$(X$, 6) + Mid$(X$, 9, 24) + Mid$(X$, 75)
             Call REPLA(Y$, Mid$(X$, 7, 2), 1, 2)
             Call REPLA(Y$, Mid$(X$, 7, 2), 45, 2)
             Call REPLA(Y$, Mid$(X$, 33, 42), 3, 42)
             Call REPLA(Y$, MKS$(PRUNI), 85, 4)
             Call REPLA(Y$, MKS$(CAPEDID), 81, 4)
             Call REPLA(Y$, MKS$(CAYEN), 97, 4)
             Call REPLA(Y$, MKS$(CASAL), 101, 4)
             Call REPLA(Y$, MKS$(CAPRE), 111, 4)
             '
             JK& = JK& + 1
             Call GRAREG("LIPEDPE", Y$, JK&)
           End If
         End If
       End If
     End If
29 Next KU&
   '
30 Call SETULTREG("LIPEDPE", JK&)
   Call CIERRARCH("LIPEDPE")
   '
99 Screen.MousePointer = 1
   Exit Sub
   '
End Sub
'
Sub CALPEPREPVIE()
   '
   Screen.MousePointer = 11
   JJ& = 0: LL& = 0
   FIN& = ULTREG&("PEDDETA")
   '
   Call HEADERS("LIPEDPE", "")
   '
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     X$ = REGLEIDO$("PEDDETA", U&)
     NC% = CVI(Mid$(X$, 7, 2))
     If Asc(Mid$(X$, 27, 1)) <> 9 Then ' no esta anulado
       If Asc(Mid$(X$, 27, 1)) <> 3 Then ' no esta entregado
         CI% = CVI(Mid$(X$, 33, 2))
         If CI% > 0 And CI% <= NUMAS% Then
           CAPEN = CVS(Mid$(X$, 79, 4)) - CVS(Mid$(X$, 95, 4))
           If CAPEN > 0.05 Then
             X1$ = X$
             Y$ = String$(46, 0) + Left$(X$, 6) + Mid$(X$, 9, 24) + Mid$(X$, 75)
             Call REPLA(Y$, Mid$(X$, 7, 2), 1, 2)
             Call REPLA(Y$, Mid$(X$, 7, 2), 45, 2)
             Call REPLA(Y$, Mid$(X$, 33, 42), 3, 42)
             Call REPLA(Y$, MKS$(CAPEN), 101, 4)
             '
             CAYAPRE = 0
             IPEDIT$ = TRIM$(Mid$(X$, 113, 10))
             If IPEDIT$ <> "" Then
               RFG$ = RECFILT$("PEDIPREP", 1, IPEDIT$)
               If Len(RFG$) >= 4 Then
                 CAYAPRE = 0
                 For KU& = 1 To Len(RFG$) Step 4
                   REPRE& = CVS(Mid$(RFG$, KU&, 4))
                   YY$ = REGLEIDO$("PEDIPREP", REPRE&)
                   CAYAPRE = CAYAPRE + CVS(Mid$(YY$, 91, 4))
                 Next KU&
               End If
             End If
             '
             Call REPLA(Y$, MKS$(CAYAPRE), 111, 4)
             CANOPRE = CAPEN - CAYAPRE
             If CANOPRE < 0 Then CANOPRE = 0
                  '
                  'PREUNI = CVS(Mid$(X$, 83, 4)) * (100 - PORDESCLI(NC%)) / 100
                  'If OPX% <> 2 Then PREUNI = 0
                  'ITOTA# = CDbl(CAPEN * PREUNI)
                  'INOAS# = CDbl(CANOPRE * PREUNI)
                  '
                  'Call REPLA(Y$, MKS$(PREUNI), 85, 4)
                  'Call REPLA(Y$, MKD$(ITOTA#), 89, 8)
                  'Call REPLA(Y$, MKS$(CANOASI), 111, 4)
                  'Call REPLA(Y$, MKD$(INOAS#), 119, 8)
                  'Call REPLA(Y$, Space$(34), 127, 34)
                  '
             JJ& = JJ& + 1
             Call GRAREG("LIPEDPE", Y$, JJ&)
           End If
         End If
       End If
     End If
   Next U&
   '
30 Call SETULTREG("LIPEDPE", JJ&)
   Call CIERRARCH("LIPEDPE")
   '
   Screen.MousePointer = 1
End Sub

Private Sub Command64_Click()
    Command64.Enabled = False
    '
    Call CALPEPREP
    Call FILESORT("LIPEDPE", "", 3, 3, "ASC")
    '
    If Option1.Value = True Then
       Sali$ = "*COPREMA"
     Else
       Sali$ = "*LIPREMA"
    End If
    Call FINAL(Sali$)
    '
    Command64.Enabled = True
End Sub

Private Sub Form_Load()
   '
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
   XX$ = FILTERGETRECORD$("PEDENCA", 1, MKS$(NPEDI&))
   If CVS(Left$(XX$, 4)) = NPEDI& Then
     NCLIE% = CVI(Mid$(XX$, 55, 2))
     If NCLIE% > 0 Then
       VDEF$ = Space$(124) + String$(4, 0) + Chr$(POSI%)
       Call REPLA(VDEF$, RASOCLI$(NCLIE%), 1, 36)
         CAENTRE$ = Mid$(XX$, 73, 32)
       Call REPLA(VDEF$, CAENTRE$, 37, 36)
         LOENTRE$ = Mid$(XX$, 105, 32)
       Call REPLA(VDEF$, LOENTRE$, 73, 36)
         NOCOM$ = Mid$(XX$, 57, 16)
       Call REPLA(VDEF$, NOCOM$, 109, 16)
       Call REPLA(VDEF$, MKI$(1), 125, 2)
       Call REPLA(VDEF$, MKI$(1), 127, 2)
       '
       KKI% = 0
       '
       ' \\\OT-5-519-FAG-15/07/05
10     DEP1% = CInt(Val(CONTROL$("", "DEP0EXPE")))
       DEP2% = CInt(Val(CONTROL$("", "DEPPREPA")))
       '
       Call REPLA(VDEF$, Chr$(DEP1%), 130, 1)
       Call REPLA(VDEF$, Chr$(DEP2%), 131, 1)
       ' \\\OT-5-519-FAG-FIN
       Call REPLA(VDEF$, Chr$(POSI%), 129, 1)
       '
       OBX$ = OBJPLA$("ETIQ-DESPA", VDEF$)
       If OBX$ <> "" Then
         'POSI% = Asc(Mid$(OBX$, 129, 1))
         FORIPRE$ = TRIM$(CONTROL$("", "ETIDESP"))
         If FORIPRE$ = "" Then FORIPRE$ = "ETIQ-EXP"
         PORTUREC$ = TRIM$(CONTROL$(IDENTER$, "PUERDES"))
         If PORTUREC$ = "" Then PORTUREC$ = CONTROL$("*", "PUERDES")
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
20         CACOLTAB1% = 0
           CODPARAM$(1) = "TRANSPOR": DOCPARAM$(1) = TRIM$(Mid$(OBX$, 1, 36))
           CODPARAM$(2) = "DOMI-TRA": DOCPARAM$(2) = TRIM$(Mid$(OBX$, 37, 36))
           CODPARAM$(3) = "LOCA-TRA": DOCPARAM$(3) = TRIM$(Mid$(OBX$, 73, 36))
           CODPARAM$(4) = "NRO-OCOM": DOCPARAM$(4) = TRIM$(Mid$(OBX$, 109, 16))
           CODPARAM$(5) = "REF-MAT1": DOCPARAM$(5) = TRIM$(Str$(CVI(Mid$(OBX$, 125, 2))))
           CODPARAM$(6) = "REF-MAT2": DOCPARAM$(6) = TRIM$(Str$(CVI(Mid$(OBX$, 127, 2))))
           '\\\OT-06-0504-RG-23/11/06///
           CODPARAM$(7) = "NRO-AENT": DOCPARAM$(7) = TRIM$(Str$(NPEDI&))
           CAPARDOC% = 7
           '\\\OT-06-0504-RG-FIN///
           '
           CANCELPRINT% = 0
           NOCONFIRM% = 1
           'For KKI% = 1 To TOBUL%
           KKI% = KKI% + 1
           CODPARAM$(5) = "REF-MAT1": DOCPARAM$(5) = TRIM$(Str$(KKI%))
           'If POSI% = 1 Then
               Call PRINTDOC("@" + FORIPRE$) ' Etiqueta de Expedición
           '  Else
           '    Call PRINTDOC("@ETIQ-EXQ")   ' Etiqueta de Expedición
           'End If
           'POSI% = (POSI% + 1) Mod 2
           'If POSI% = 0 Then POSI% = 2
           If KKI% < TOBUL% Then GoTo 20
           'Next KKI%
           '
         End If
         Call SETPRINPORT("RESTORE")
       End If
     End If
   End If
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


