VERSION 5.00
Begin VB.Form MOVISTK04 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0EEC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stocks - Movimientos"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5820
   Icon            =   "MOVISTK04.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   5820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command8 
      Enabled         =   0   'False
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
      Left            =   945
      Picture         =   "MOVISTK04.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Remito de Traslado"
      Top             =   5460
      Width           =   560
   End
   Begin VB.CommandButton Command10 
      Enabled         =   0   'False
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
      Left            =   105
      Picture         =   "MOVISTK04.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Lista de Empaque"
      Top             =   5460
      Width           =   560
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00004000&
      Height          =   5370
      Left            =   4935
      ScaleHeight     =   5310
      ScaleWidth      =   1005
      TabIndex        =   19
      Top             =   0
      Width           =   1065
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   23
         Top             =   3885
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   24
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command1 
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
         Picture         =   "MOVISTK04.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   200
         Width           =   650
      End
      Begin VB.CommandButton Command29 
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
         Picture         =   "MOVISTK04.frx":0A68
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   890
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
         Height          =   690
         Left            =   105
         Picture         =   "MOVISTK04.frx":0D72
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1815
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "MOVISTK04.frx":11B4
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   4280
         Width           =   1515
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0EEC0&
      Caption         =   "Movimientos de Stock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4600
      Left            =   210
      TabIndex        =   3
      Top             =   105
      Width           =   4530
      Begin VB.Frame Frame8 
         BackColor       =   &H00C0EEC0&
         Caption         =   "Expedición y Despacho"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   210
         TabIndex        =   6
         Top             =   1785
         Width           =   4110
         Begin VB.PictureBox Picture8 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   735
            ScaleHeight     =   585
            ScaleWidth      =   2535
            TabIndex        =   11
            Top             =   420
            Width           =   2600
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
               Height          =   600
               Left            =   0
               Picture         =   "MOVISTK04.frx":30D6
               Style           =   1  'Graphical
               TabIndex        =   27
               ToolTipText     =   "Expedición - Remito de Ventas - Clientes "
               Top             =   0
               Width           =   560
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
               Height          =   600
               Left            =   1995
               Picture         =   "MOVISTK04.frx":33E0
               Style           =   1  'Graphical
               TabIndex        =   17
               ToolTipText     =   "Remito en Consignacion - Proveedores de Servicios"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label8 
               BackStyle       =   0  'Transparent
               Caption         =   "Comprobantes de  Despacho"
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
               Left            =   630
               TabIndex        =   18
               Top             =   105
               Width           =   1215
            End
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0EEC0&
         Caption         =   "Fábrica - Ajustes - Transferencias"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   210
         TabIndex        =   5
         Top             =   3150
         Width           =   4110
         Begin VB.PictureBox Picture10 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   210
            ScaleHeight     =   585
            ScaleWidth      =   3630
            TabIndex        =   12
            Top             =   420
            Width           =   3690
            Begin VB.CommandButton Command17 
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
               Left            =   3070
               Picture         =   "MOVISTK04.frx":36EA
               Style           =   1  'Graphical
               TabIndex        =   16
               ToolTipText     =   "Descartes de Stock"
               Top             =   0
               Width           =   560
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
               Height          =   600
               Left            =   560
               Picture         =   "MOVISTK04.frx":39F4
               Style           =   1  'Graphical
               TabIndex        =   15
               ToolTipText     =   "Transferencias de Stock"
               Top             =   0
               Width           =   560
            End
            Begin VB.CommandButton Command25 
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
               Picture         =   "MOVISTK04.frx":3CFE
               Style           =   1  'Graphical
               TabIndex        =   14
               ToolTipText     =   "Ajuste de Existencias"
               Top             =   0
               Width           =   560
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
               Height          =   600
               Left            =   2520
               Picture         =   "MOVISTK04.frx":3E48
               Style           =   1  'Graphical
               TabIndex        =   2
               ToolTipText     =   "Vale de Consumo de Materiales"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "Movimientos     Varios"
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
               Left            =   1260
               TabIndex        =   13
               Top             =   105
               Width           =   1110
            End
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00C0EEC0&
         Caption         =   "Entradas a Depósito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   210
         TabIndex        =   4
         Top             =   420
         Width           =   4110
         Begin VB.PictureBox Picture1 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   2205
            ScaleHeight     =   585
            ScaleWidth      =   1635
            TabIndex        =   9
            Top             =   420
            Width           =   1695
            Begin VB.CommandButton Command18 
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
               Picture         =   "MOVISTK04.frx":3F92
               Style           =   1  'Graphical
               TabIndex        =   1
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Propia Producción"
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
               Left            =   630
               TabIndex        =   10
               Top             =   105
               Width           =   1110
            End
         End
         Begin VB.PictureBox Picture6 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   210
            ScaleHeight     =   585
            ScaleWidth      =   1740
            TabIndex        =   7
            Top             =   420
            Width           =   1800
            Begin VB.CommandButton Command28 
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
               Picture         =   "MOVISTK04.frx":429C
               Style           =   1  'Graphical
               TabIndex        =   0
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Recepción Proveedores"
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
               Left            =   630
               TabIndex        =   8
               Top             =   105
               Width           =   1110
            End
         End
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "MOVISTK04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Call CIERRARCH("*.*")
   Hide
End Sub

Private Sub Command10_Click()
    Command10.Enabled = False
    Call CONECRUN("LISEMPAQ", "Generación de Listas de Empaque")
    Command10.Enabled = True
End Sub


Sub Command17_Click()
    Command17.Enabled = False
    If DERACCE%("DESCRAP", "Descartes de Stock") >= 2 Then
      Call DESCRAP
    End If
99  Command17.Enabled = True
End Sub

Sub Command18_Click()
    Call RECPROP
End Sub

Sub Command21_Click()
    Command21.Enabled = False
    If DERACCE%("VALEMAT", "Vale de Consumo de Materiales") >= 2 Then
       Call VALEMAT
    End If
    Command21.Enabled = True
End Sub

Sub Command22_Click()
    Call NUENTRECONSI(NPROVE%)
End Sub

Sub Command23_Click()
    Command23.Enabled = False
    MODU$ = "REMITO04"
    Call CONECRUN(MODU$, "Emision de Remito de Facturacion")
    Command23.Enabled = True
End Sub

Sub Command24_Click()
   Command24.Enabled = False
   If DERACCE%("TRASTOCK", "Transferencias de Stock") >= 2 Then
      TRANSTK04.Show 1
   End If
   Command24.Enabled = True
End Sub

Sub Command25_Click()
    Command25.Enabled = False
    If DERACCE%("AJUSTOCK", "Ajustes de Stock") >= 2 Then
       PREPA% = 0
       FORMAJU% = 0
       AJUSTK04.Show 1
    End If
    Command25.Enabled = True
End Sub

Sub Command26_Click()
'   Call CONECRUN("FSETUP04/OPTRAJUST", "Configuración de Funcionamiento")
   OPTRAJU04.Show 1
   Call FINAL("")
End Sub

Sub Command28_Click()
    Command28.Visible = False
    ACONEC$ = "BRECEP04"
    Call CONECRUN(ACONEC$, "Recepcion de Materiales")
    Command28.Visible = True
End Sub

Private Sub Command29_Click()
    Call HELPONLINE("CONSTOCK")
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    If EXISTE%("RETRAS02.EXE") > 0 Then
        Call CONECRUN("RETRAS02", "Emision de Remito de Traslado")
      Else
        Call CONECRUN("REMITRAS", "Emision de Remito de Traslado")
    End If
    Command8.Enabled = True
End Sub

Private Sub Form_Activate()
    TEXPE$ = "Remito de Ventas - Clientes"
    If CONTROL$("", "MODOEXPE") = "ODESPACH" Then
        TEXPE$ = "Orden de Despacho"
    End If
    Command23.ToolTipText = TEXPE$
    Call FRESHALL
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     OPENFORMS = DoEvents
   Wend
End Sub

Private Sub Form_Load()
    '
    Call TRANSLABELS(Name)
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    Call GRATITFOR(Caption)
    '
    If EXISTE%("REMITO04.EXE") > 0 Then Command23.Enabled = True
    If EXISTE%("ORCOMP04.EXE") > 0 Then Command22.Enabled = True
    If EXISTE%("RETRAS04.EXE") > 0 Then Command8.Enabled = True
    If EXISTE%("LISEMPAQ.EXE") > 0 Then Command10.Enabled = True
    '
End Sub

Private Sub Form_Resize()
   If WindowState = 1 Then
      Call CIERRARCH("*.*")
      Call BAJALDISCO
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Hide
End Sub

Sub RECPROP()
    '
    If PRECOCIE% = 0 Then
       PRECOCIE% = (-1)
       If CONTROL$("", "PRECOCIE") = "SI" Then
          PRECOCIE% = 1
       End If
    End If
    '
10  HOOI% = HOY(HO$)
    If ULTREG&("ORDEFABR") < 1 Then   ' por numero de O/F
        GoTo 100
    End If
    '
    modo% = COMALTER%("Puede Declarar Entradas a Stock:\  \a) por Cierre Parcial o Total de Orden de Fabricación\o\b) en Forma Directa por Armado de Conjunto\\Cierre de O./Fabricación\Armado de Conjunto\Cancelar")
    On modo% GoTo 20, 100
    Exit Sub
    '
20  Call ORFCIE
    Exit Sub
    '
100 Screen.MousePointer = 11
    For U& = ULTREG&("RECPROP") To 1 Step -1
        XX$ = REGLEIDO$("RECPROP", U&)
        If CVS(Mid$(XX$, 59, 4)) > 0 Then
            UPARTE& = CVS(Mid$(XX$, 59, 4))
            GoTo 105
        End If
    Next U&
    UPARTE& = 0
    '
105 Screen.MousePointer = 1
    Call BLANARCH("!CAREPRO")
    '
    VDEF$ = String$(88, 0)
    Call REPLA(VDEF$, MKI$(HOOI%), 1, 2)
    Call REPLA(VDEF$, MKS$(UPARTE& + 1), 85, 4)
    Call REPLA(VDEF$, Chr$(0), 37, 1)
    Call REPLA(VDEF$, Chr$(0), 38, 1)
    '
110 XX$ = OBJPLA$("RECPROP", VDEF$)
    If XX$ = "" Then Exit Sub
    '
160 VDEF$ = XX$
    FECHA% = CVI(Left$(XX$, 2))
    If FECHA% > HOOI% Or FECHA% < 1 Then
        Call MENSERR(24, "Fecha Incorrecta")
        GoTo 110
    End If
    '
'    CI% = CVI(Mid$(XX$, 3, 2))
'    If TRIM$(CODEXT$(CI%)) = "" Then
'        Call MENSERR(24, "Codigo Inexistente.")
'        GoTo 110
'    End If
'    CODCON$ = CODEXT$(CI%)
'    '
    DP1% = Asc(Mid$(XX$, 37, 1))
    DP2% = Asc(Mid$(XX$, 38, 1))
    If TRIM$(ECOARCH$("ECODEP", Chr$(DP1%))) = "" Then
        Call MENSERR(24, "Falta Depósito de Entrada")
        GoTo 110
    End If
    '
    If TRIM$(ECOARCH$("ECODEP", Chr$(DP2%))) = "" Then
        Call MENSERR(24, "Falta Depósito de Consumo")
        GoTo 110
    End If
    '
120 U3& = 0
    For U1& = 1 To ULTREG&("!CAREPRO")
      XY$ = REGLEIDO$("!CAREPRO", U1&)
      CANTI = CVD(Mid$(XY$, 23, 8))
      If Abs(CANTI) >= 0.05 Then
        U3& = U3& + 1
        Call GRAREG("!CAREPRO", XY$, U3&)
      End If
    Next U1&
    Call SETULTREG("!CAREPRO", U3&)
    '
    VTB% = VENTABU%("CAREPRO")
    If VTB% < 0 Then GoTo 110
    '
    CARE& = ULTREG&("!CAREPRO")
    If CARE& < 1 Then GoTo 120
    '
    NUORIT% = 0
    For U& = 1 To CARE&
          KKL& = 0
       On Error GoTo RECUERR
       XY$ = REGLEIDO$("!CAREPRO", U&)
          KKL& = 10
       CANTI = CVD(Mid$(XY$, 23, 8))
       SCRAP = 0
       CI% = CVI(Left$(XY$, 2))
       CODICONJ$ = CODEXT$(CI%)
       CICONJ% = CI%
          KKL& = 20
       '
       '\\\OT-06-0165-WAR-08/05/06///
       'If CANTI < 0 Or SCRAP < 0 Or CANTI + SCRAP < 0.05 Then
       If SCRAP < 0 Then
       '\\\OT-06-0165-WAR-FIN///
          Call MENSERR(24, "Cantidades Incorrectas en Código " + TRIM$(CODICONJ$))
          GoTo 120
       End If
          KKL& = 30
       '
       NORFABRI$ = ""
       CACONJU# = CDbl(CANTI + SCRAP)
       VACONTROL% = 1
          KKL& = 40
       '
       Call SETULTREG("!MATEROF", 0)
       Call CIERRARCH("!MATEROF")
          KKL& = 50
       Call CARCOMPO(CODICONJ$, NORFABRI$, CACONJU#, VACONTROL%, 2)
          KKL& = 60
       '
       If VACONTROL% < 0 Then GoTo 120           ' cancelacion
       '
       '\\\OT-05-0730-WAR-10/11/05///
       CAITMO% = 0
          KKL& = 70
         For YX& = 1 To ULTREG&("!MATEROF")
            YY$ = REGLEIDO$("!MATEROF", YX&)
            CI% = CVI(Mid$(YY$, 1, 2))
            CAN# = CVS(Mid$(YY$, 5, 4))
             If CI% > 0 Then
               If CI% <= NUMAS% Then
                 If CAN# > 0.005 Then
                   CAITMO% = CAITMO% + 1
                   CIMOVIM%(CAITMO%) = CI%
                   CAMOVIM(CAITMO) = (-1) * CAN#
                 End If
               End If
             End If
         Next YX&
          KKL& = 80
       If PUEDEMOVI%(DP2%) < 0 Then GoTo 110 ' CANCELA SALIDA
          KKL& = 90
       '\\\OT-05-0730-WAR-FIN///
       '
       NRO& = CVS(Mid$(XX$, 85, 4))
       NROOF& = CVS(Mid$(XY$, 19, 4))
       If NROOF& < 1 Then NROOF& = NRO&
          KKL& = 100
       DOPRI$ = "OF-" + TRIM$(Str$(NROOF&))
       DOSEC$ = "PF-" + TRIM$(Str$(NRO&))
       LOTE$ = Mid$(XY$, 3, 16)
       If LOTE$ = "" Then
          LOTE$ = DOPRI$
       End If
          KKL& = 110
       REFE$ = Mid$(XX$, 55, 30)
       If TRIM$(REFE$) = "" Then
          REFE$ = "Informe de Produccion"
       End If
       If CANTI < (-0.005) Then
          REFE$ = "Desarmado de Conjunto"
       End If
          KKL& = 120
       '
       NC% = 0
       VUNI# = 0
       MONE$ = ""
       MONEII% = 0
       CMOV$ = "PF"
       '
       NUORIT% = 0
       DEPX% = DP1%: If DP1% = 99 Then DEPX% = DEPOPRE%(CICONJ%)
          KKL& = 130
       Call MOVISTOK(FECHA%, CICONJ%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NUORIT%, VUNI#, MONEII%, NC%, DEPX%, CANTI)
          KKL& = 140
       '
       YY$ = REGBLAN$("RECPROP")
       Call REPLA(YY$, MKI$(FECHA%), 1, 2)
       Call REPLA(YY$, Chr$(DP1%), 3, 1)
       Call REPLA(YY$, Chr$(DP2%), 4, 1)
       Call REPLA(YY$, REFE$, 5, 32)
       Call REPLA(YY$, MKI$(CI%), 37, 2)
       Call REPLA(YY$, MKS$(CANTI), 39, 4)
       Call REPLA(YY$, LOTE$, 43, 16)
       Call REPLA(YY$, MKS$(NRO&), 59, 4)
       Call REPLA(YY$, MKI$(-1), 63, 2)
          KKL& = 150
       Call REGAPP("RECPROP", YY$)
          KKL& = 160
       COD$ = CODEXT$(CICONJ%)
       CODCON$ = COD$
       '
          KKL& = 170
       For U2& = 1 To ULTREG&("!MATEROF")
          ZZ$ = REGLEIDO$("!MATEROF", U2&)
          CI2% = CVI(Left$(ZZ$, 2))
          CAN = -1 * CVS(Mid$(ZZ$, 5, 4))
          '
          If TRIM$(UCase$(UNIMED$(CI2%))) = "GK" Then
             CAN = CAN / 1000
          End If
          '
          If CI2% > 0 Then
             If CI2% <= NUMAS% Then
                If Abs(CAN) >= 0.05 Then
                   LOTE$ = ""
                   CMOV$ = "CF"
                   DOPRI$ = "OF-" + TRIM$(Str$(NROOF&))
                   DOSEC$ = "PF-" + TRIM$(Str$(NRO&))
                   REFE$ = "CONSUMO S/DESP." + CODCON$
                   '\\\OT-06-0165-WAR-08/05/06///
                   If CAN > 0 Then
                    REFE$ = "DESARMADO CONJ." + CODCON$
                   End If
                   '\\\OT-06-0165-WAR-FIN///
                   DEPX% = DP2%: If DP2% = 99 Then DEPX% = DEPOPRE%(CI2%)
                   '
                   If STODEPOS(CI2%, DEPX%) < (-1) * CAN Then
                     If InStr(EPAC$, "YALTRE") > 0 Then
                       Call CARDEPOS
                       For KKII% = 1 To NUSU%
                         If STODEPOS(CI2%, COSU%(KKII%)) >= (-1) * CAN Then
                           DEPX% = COSU%(KKII%)
                           GoTo 124
                         End If
                       Next KKII%
                     End If
                   End If
                   '
124                NUORIT% = NUORIT% + 1
                   Call MOVISTOK(FECHA%, CI2%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NUORIT%, VUNI#, MONEII%, NC%, DEPX%, CAN)
                End If
             End If
          End If
       Next U2&
          KKL& = 180
       Call CIERRARCH("!MATEROF")
       '
       XY$ = REGLEIDO$("!CAREPRO", U&)
       Call REPLA(XY$, MKD$(0), 23, 8)
       Call GRAREG("!CAREPRO", XY$, U&)
          KKL& = 190
       '
       On Error GoTo 0
129 Next U&
    Call CIERRARCH("MOVISTO")
    Call CIERRARCH("RECPROP")
    '
180 Call CIERRARCH("*.*")
    Call BLANARCH("!CAREPRO")
    Screen.MousePointer = 1
    GoTo 100
    '
RECUERR:
    MENSA$ = "Se ha Producido un Error " & Err.Number & " al Procesar el Código '" & TRIM$(CODICONJ$) & "' en el Registro Número " & U& & " y la Línea Número " & KKL& & "."
    MENSA$ = MENSA$ + " Informe a FLEXOFT para Recibir Orientación Correctiva: 4524-1284 / 4522-2048 (Int.32)"
    MsgBox MENSA$, vbExclamation
    Call CIERRARCH("*.*")
    Call FINAL("")
    '
End Sub

Private Sub Timer1_Timer()
    '
End Sub
