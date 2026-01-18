VERSION 5.00
Begin VB.Form OPREMSTAN 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Remitos y Ordenes de Despacho - Opciones"
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5160
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command11 
      Caption         =   "Cumplim"
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
      Height          =   855
      Left            =   4080
      Picture         =   "OPREMSTAN.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Indice de Cumplimiento de Entregas"
      Top             =   4575
      Width           =   855
   End
   Begin VB.CommandButton Command19 
      Caption         =   "N.Orig"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4095
      Picture         =   "OPREMSTAN.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Códigos Equivalentes de Clientes"
      Top             =   3555
      Width           =   855
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   4095
      ScaleHeight     =   75
      ScaleWidth      =   810
      TabIndex        =   20
      Top             =   2835
      Width           =   870
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Resúmenes de  Entregas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Index           =   2
      Left            =   210
      TabIndex        =   10
      Top             =   2100
      Width           =   3690
      Begin VB.PictureBox Picture8 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   955
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   32
         Top             =   2520
         Width           =   2430
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
            Height          =   600
            Left            =   0
            Picture         =   "OPREMSTAN.frx":0A14
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Ultima Entrega por Producto"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Ultima Entrega por Producto"
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
            Left            =   735
            TabIndex        =   33
            Top             =   60
            Width           =   1515
         End
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   795
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   23
         Top             =   1995
         Width           =   2430
         Begin VB.CommandButton Command8 
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
            Picture         =   "OPREMSTAN.frx":0D1E
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Consulta / Listado de Material en Consignación (Clientes)"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Listado Material en Consignacion"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   750
            Left            =   735
            TabIndex        =   25
            Top             =   55
            Width           =   1500
         End
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   635
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   17
         Top             =   1470
         Width           =   2430
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
            Height          =   600
            Left            =   0
            Picture         =   "OPREMSTAN.frx":1028
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Listado Secuencial por Número de Comprobante"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Secuencial por Nro. de Comprobante"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   750
            Left            =   630
            TabIndex        =   19
            Top             =   55
            Width           =   1750
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   475
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   11
         Top             =   945
         Width           =   2430
         Begin VB.CommandButton Command3 
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
            Left            =   -20
            Picture         =   "OPREMSTAN.frx":1332
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Listado de Entregas por Cliente "
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Resumen de Entre- gas por Cliente"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   750
            Left            =   630
            TabIndex        =   13
            Top             =   55
            Width           =   1695
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   315
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   14
         Top             =   420
         Width           =   2430
         Begin VB.CommandButton Command4 
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
            Picture         =   "OPREMSTAN.frx":1774
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Listado de Entregas por Artículo "
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Resumen de Entre- gas por Producto"
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
            Left            =   630
            TabIndex        =   16
            Top             =   55
            Width           =   1725
         End
      End
   End
   Begin VB.CommandButton Command2 
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
      Height          =   855
      Left            =   4095
      Picture         =   "OPREMSTAN.frx":1A7E
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Terminar - Salir de la Aplicación"
      Top             =   315
      Width           =   855
   End
   Begin VB.CommandButton Command7 
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
      Height          =   855
      Left            =   4095
      Picture         =   "OPREMSTAN.frx":1BC8
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
      Top             =   1260
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Setup - Configuración"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Index           =   1
      Left            =   210
      TabIndex        =   4
      Top             =   5565
      Width           =   3690
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   630
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   5
         Top             =   420
         Width           =   2430
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
            Height          =   600
            Left            =   0
            Picture         =   "OPREMSTAN.frx":1ED2
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Acceso Directo a Configuración de Funcionaliento"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Configuracion de Funcionamiento"
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
            TabIndex        =   7
            Top             =   105
            Width           =   1635
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Operaciones Complementarias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Index           =   0
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   3690
      Begin VB.PictureBox Picture7 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   630
         ScaleHeight     =   585
         ScaleWidth      =   2685
         TabIndex        =   27
         Top             =   945
         Width           =   2745
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
            Height          =   600
            Left            =   0
            Picture         =   "OPREMSTAN.frx":201C
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "Devolución Material en Consignación"
            Top             =   0
            Width           =   560
         End
         Begin VB.CommandButton Command10 
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
            Left            =   550
            Picture         =   "OPREMSTAN.frx":2166
            Style           =   1  'Graphical
            TabIndex        =   28
            ToolTipText     =   "Ingreso por Devolución de Clientes "
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Devolución de Materiales"
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
            Left            =   1365
            TabIndex        =   29
            Top             =   90
            Width           =   1200
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   315
         ScaleHeight     =   585
         ScaleWidth      =   2700
         TabIndex        =   1
         Top             =   420
         Width           =   2760
         Begin VB.CommandButton Command6 
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
            Left            =   550
            Picture         =   "OPREMSTAN.frx":22B0
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Eliminación Solo de Lista de Pendientes de Factura"
            Top             =   0
            Width           =   570
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
            Picture         =   "OPREMSTAN.frx":25BA
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Anulación General de Remito / Orden de Despacho Emitidos"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Anulación de Comprobantes"
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
            Left            =   1320
            TabIndex        =   3
            ToolTipText     =   "Click Aqui Importa Archivo GETRAMO"
            Top             =   50
            Width           =   1740
         End
      End
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   3555
      Picture         =   "OPREMSTAN.frx":28C4
      Top             =   6300
      Width           =   2010
   End
End
Attribute VB_Name = "OPREMSTAN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    '
    Command1.Enabled = False
    OPREMIFAC.Show 1
    Command1.Enabled = True
    '
End Sub

Private Sub Command10_Click()
   Command10.Enabled = False
10 HOII% = HOY(HOS$)
   VDEF$ = REGBLAN$("MOVISTO")
   '
20 Call REPLA(VDEF$, MKI$(HOII%), 1, 2)
   OBX$ = OBJPLA$("DEVOLUMAT", VDEF$)
   If OBX$ <> "" Then
     VDEF$ = OBX$
     FF% = CVI(Mid$(OBX$, 1, 2))
     If FF% < 1 Or FF% > HOII% Then
       Call MENSERR(24, "Fecha no Válida")
       GoTo 20
     End If
     '
     CIIX% = CVI(Mid$(OBX$, 3, 2))
     If CIIX% < 1 Or CIIX% > NUMAS% Then
       Call MENSERR(24, "Codigo no Válido")
       GoTo 20
     End If
     '
     LOTE$ = TRIM$(Mid$(OBX$, 5, 16))
     CMOV$ = "RT"
     '
     DOPRI$ = TRIM$(Mid$(OBX$, 23, 10))
     If DOPRI$ = "" Then
       Call MENSERR(24, "Falta Comprobante del Movimiento")
       GoTo 20
     End If
     DOSEC$ = "DV-" + DOPRI$
     '
     NC% = CVI(Mid$(OBX$, 84, 2))
     If NC% < 1 Or ECOARCH$("DEUDOR1", MKI$(NC%)) = "" Then
       Call MENSERR(24, "Cliente Inexistente")
       GoTo 20
     End If
     REFE$ = RASOCLI$(NC%)
     '
     CANTI = CVD(Mid$(OBX$, 88, 8))
     If CANTI < 0.05 Then
       Call MENSERR(24, "Falta Cantidad Devuelta")
       GoTo 20
     End If
     '
     DEPO% = DEPOEXPE%
     '
     Call MOVISTO(FF%, CIIX%, LOTE$, CMOV$, DOPRI$, DOSEC$, REFE$, VUNI#, MONE$, NC%, DEPO%, CANTI)
     Call CIERRARCH("MOVISTO")
     Call COMUNI("Devolución Registrada Correctamente !!!")
     GoTo 10
   End If
   '
   Command10.Enabled = True
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   Call CIERRARCH("*.*")
   Call CONECRUN("COCUMENT", "Control de Cumplimiento de Entregas")
   Command11.Enabled = True
End Sub

Private Sub CONTROFAC()
    '
10  FEX = HOY(HO$)
    HAS% = FEX
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    DES% = FECHANUM(DESS$)
    '
20  VDEF$ = MKI$(DES%) + MKI$(HAS%)
    '
    OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
        GoTo 99
    End If
    '
    DES% = CVI(Left$(OBX$, 2))
    HAS% = CVI(Mid$(OBX$, 3, 2))
    If DES% > HAS% Then
        GoTo 20
    End If
    FEX = DES%: PERIO$ = FECHATEX$(FEX)
    FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    '
    Screen.MousePointer = 11
    J& = 0
    '
    FIN& = ULTREG&("MOVISTO")
    Screen.MousePointer = 11
    TOTAGEN# = 0
    For U& = 1 To FIN&
        '
        Call TRACE(20, U&, FIN&)
        X$ = REGLEIDO$("MOVISTO", U&)
        CLI% = CVI(Mid$(X$, 84, 2))
        If CLI% > 0 Then
           FF% = CVI(Left$(X$, 2))
           If FF% >= DES% Then
              If FF% <= HAS% Then
                 CII% = CVI(Mid$(X$, 3, 2))
                 If CII% > 0 Then
                    If CII% <= NUMAS% Then
                       If ECOARCH$("DEUDOR1", MKI$(CLI%)) <> "" Then
                          '
                          REMIFAC$ = Mid$(X$, 23, 20)
                          CANTU# = CVD(Mid$(X$, 96, 8)) - CVD(Mid$(X$, 88, 8))
                          '
                          Z$ = REGBLAN$("FICHENT")
                          Call REPLA(Z$, MKI$(CLI%), 1, 2)
                          'Call REPLA(Z$, PERIO$, 3, 20)
                          Call REPLA(Z$, MKI$(FF%), 23, 2)
                          Call REPLA(Z$, MKI$(CII%), 25, 2)
                          Call REPLA(Z$, MKD$(CANTU#), 27, 8)
                          Call REPLA(Z$, REMIFAC$, 35, 20)
                          '
                          J& = J& + 1
                          Call GRAREG("FICHENT", Z$, J&)
                          '
                       End If
                    End If
                 End If
              End If
           End If
        End If
    Next U&
    '
    Call SETULTREG("FICHENT", J&)
    Call CIERRARCH("FICHENT")
    Screen.MousePointer = 1
    '
    If J& > 0 Then
      Call HEADERS("FICHENT", "")
      Call HEADERS("FICHENT", "Periodo: " + PERIO$)
      '
      Call FILESORT("FICHENT", "FICHENT", 1, 1, "ASC")
      '
      Call FINAL("*LIFICHEN")
    End If
    '
99  Call CIERRARCH("*.*")
    '
End Sub

Private Sub Command12_Click()
  '\\\OT-06-0070-WAR-01/03/06///
  Dim FUE%(32767), NCLIX%(32767)
  Dim RTO$(32767), FACX$(32767)
  Dim QT#(32767)
  '
  Command12.Enabled = False
  Call DESHASFECHA(DES%, HAS%, PERIO$)
  '
  Screen.MousePointer = 11
  FIN& = ULTREG&("MOVISTO")
  For J& = 1 To FIN&
    Call TRACE(20, J&, FIN&)
    X$ = REGLEIDO$("MOVISTO", J&)
    CM1$ = Mid$(X$, 21, 2)
    FEC1% = CVI(Mid$(X$, 1, 2))
    CI% = CVI(Mid$(X$, 3, 2))
    If CM1$ = "RT" Then
      If FEC1% >= DES% Then
        If FEC1% <= HAS% Then
          If FEC1% >= FUE%(CI%) Then
            QE1 = CVD(Mid$(X$, 96, 8))
            NC1% = CVI(Mid$(X$, 84, 2))
            RT1$ = Mid$(X$, 23, 10)
            FC1$ = Mid$(X$, 33, 10)
            '
            FUE%(CI%) = FEC1%
            RTO$(CI%) = RT1$
            FACX$(CI%) = FC1$
            QT#(CI%) = QE1
            NCLIX%(CI%) = NC1%
            '
          End If
        End If
      End If
    End If
    '
  Next J&
  '
  JJ& = 0
  'Call BLANARCH("ULTIENT")
  For J1% = 1 To 32767
    If QT#(J1%) > 0.005 Then 'Si tuvo movimiento en el rango
      Y$ = REGBLAN$("ULTIENT")
      Call REPLA(Y$, MKI$(J1%), 1, 2)
      Call REPLA(Y$, MKI$(FUE%(J1%)), 3, 2)
      Call REPLA(Y$, MKD$(QT#(J1%)), 5, 8)
      Call REPLA(Y$, RTO$(J1%), 13, 10)
      Call REPLA(Y$, FACX$(J1%), 23, 10)
      Call REPLA(Y$, MKI$(NCLIX%(J1%)), 33, 2)
      JJ& = JJ& + 1
      Call GRAREG("ULTIENT", Y$, JJ&)
    End If
    J2& = J1%
    Call TRACE(20, J2&, 32767)
    If J1% = 32767 Then Exit For
  Next J1%
  '
  Call SETULTREG("ULTIENT", JJ&)
  Screen.MousePointer = 1
  '
  If JJ& = 0 Then
    Call COMUNI("No Existen Entregas en el Rango de Fechas Seleccionado")
    GoTo 99
  End If
  '
  Call FILESORT("ULTIENT", "ULTIENT", 1, 1, "ASC")
  '
  Call HEADERS("ULTIENT", "")
  Call HEADERS("ULTIENT", "Período: " + PERIO$)
  Call FINAL("*ULTIENT")
99 Command12.Enabled = True
  '
  '\\\OT-06-0070-WAR-FIN////
End Sub

Private Sub Command19_Click()
   Command19.Enabled = False
   Call FINAL("*CODORCL")
   Command19.Enabled = True
End Sub

Private Sub Command2_Click()
    '
    Call CIERRARCH("*.*")
    Hide
    '
End Sub

Private Sub Command25_Click()
    Command25.Enabled = False
    Dim CISL%(2048), CANSL#(2048)
    '
    CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
    If Len(CMO$) <> 2 Then
       CMO$ = "RT"
    End If
    '
    Call CARSELMESA
    '
10  Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) = "" Then
       GoTo 99
    End If
    CORRESQ$ = MESANO$(MESASEL$, DES%, HAS%)
    '
    Call FRESHECHO("SELREMI")
    Screen.MousePointer = 11
    FIN& = ULTREG&("MOVISTO")
    J& = 0: REMIA$ = ""
    '
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = REGLEIDO$("MOVISTO", U&)
       FF% = CVI(Left$(X$, 2))
       If FF% >= DES% Then
          If FF% <= HAS% Then
             REMI$ = UCase$(Mid$(X$, 23, 10))
             If Val(Mid$(REMI$, 4)) > 0 Then
               If REMI$ <> REMIA$ Then
                 REMIA$ = REMI$
                 TIMO1$ = Mid$(X$, 21, 2)
                 If TIMO1$ = CMO$ Then
                   If UCase$(Mid$(X$, 33, 2)) <> "FC" Then 'excluye facturados
                     CLI% = CVI(Mid$(X$, 84, 2))
                     If CLI% > 0 Then
                       Z$ = Space$(40) + String$(8, 0)
                       Call REPLA(Z$, REMI$, 1, 10)
                       Call REPLA(Z$, RASOCLI$(CLI%), 11, 30)
                       Call REPLA(Z$, MKS$(U&), 41, 4)
                       Call REPLA(Z$, MKI$(FF%), 45, 2)
                       For KK& = 1 To J&
                         YZ$ = REGLEIDO$("SELREMI", KK&)
                         If Left$(YZ$, 40) = Left$(Z$, 40) Then GoTo 14
                       Next KK&
                       J& = J& + 1
                       Call GRAREG("SELREMI", Z$, J&)
                     End If
                   End If
                 End If
               End If
             End If
          End If
       End If
14  Next U&
    '
    FIN& = ULTREG&("MOVIREPA")
    REMIA$ = ""
    '
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = REGLEIDO$("MOVIREPA", U&)
       FF% = CVI(Left$(X$, 2))
       If FF% >= DES% Then
          If FF% <= HAS% Then
             REMI$ = UCase$(Mid$(X$, 23, 10))
             If Val(Mid$(REMI$, 4)) > 0 Then
               If REMI$ <> REMIA$ Then
                 REMIA$ = REMI$
                 TIMO1$ = Mid$(X$, 21, 2)
                 If TIMO1$ = "RP" Then
                   If UCase$(Mid$(X$, 33, 2)) <> "FC" Then 'excluye facturados
                     CLI% = CVI(Mid$(X$, 84, 2))
                     If CLI% > 0 Then
                       Z$ = Space$(40) + String$(8, 0)
                       Call REPLA(Z$, REMI$, 1, 10)
                       Call REPLA(Z$, RASOCLI$(CLI%), 11, 30)
                       Call REPLA(Z$, MKS$((-1) * U&), 41, 4)
                       Call REPLA(Z$, MKI$(FF%), 45, 2)
                       For KK& = 1 To J&
                         YZ$ = REGLEIDO$("SELREMI", KK&)
                         If Left$(YZ$, 40) = Left$(Z$, 40) Then GoTo 24
                       Next KK&
                       J& = J& + 1
                       Call GRAREG("SELREMI", Z$, J&)
                     End If
                   End If
                 End If
               End If
             End If
          End If
       End If
24  Next U&
    '
    Call SETULTREG("SELREMI", J&)
    Call CIERRARCH("SELREMI")
    Screen.MousePointer = 1
    '
    If J& < 1 Then
       Call MENSERR(24, "No se Encontraron Remitos Posibles de Anular\en el Período Elegido.")
       GoTo 10
    End If
    '
    Call SELECHO("SELREMI")
    REMISEL$ = VALACT1$("SELREMI")
    If TRIM$(REMISEL$) = "" Then
       GoTo 10
    End If
    '
    RERE$ = REGACT$("SELREMI")
    PRIREG& = CVS(Mid$(RERE$, 41, 4))
    CLIE$ = Mid$(RERE$, 11, 30)
    FEX = CVI(Mid$(RERE$, 45, 2))
    '
    If Abs(PRIREG&) < 1 Then GoTo 10
    '
    TXT$ = "Remito Número " + TRIM$(REMISEL$) + "\Cliente: " + TRIM$(CLIE$) + "\Fecha: " + FECHATEX$(FEX) + "\   \\Abandonar\Anular"
    ALTE% = COMALTER%(TXT$)
    If ALTE% <> 2 Then GoTo 10
    '
    Call BLOQARCH("MOVISTO")
    Call BLOQARCH("MOVIREPA")
    Call BLOQARCH("REMPEN")
    '
    Screen.MousePointer = 11
    AMOVI$ = "MOVISTO"
    If PRIREG& < 0 Then
       AMOVI$ = "MOVIREPA"
    End If
    FIN& = ULTREG&(AMOVI$)
    KKJ% = 0: FECHA% = 0
    For U& = Abs(PRIREG&) To FIN&
    Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$(AMOVI$, U&)
      If TRIM$(Mid$(X$, 23, 10)) = REMISEL$ Then
        If FECHA% = 0 Then
           FECHA% = CVI(Left$(X$, 2))
           CLI% = CVI(Mid$(X$, 84, 2))
           DEPO% = Asc(Mid$(X$, 83, 1))
        End If
        CII% = CVI(Mid$(X$, 3, 2))
        CANTU# = CVD(Mid$(X$, 88, 8)) - CVD(Mid$(X$, 96, 8))
        For KKI% = 1 To KKJ%
           If CISL%(KKI%) = CII% Then
              CANSL#(KKI%) = CANSL#(KKI%) + CANTU#
              GoTo 20
           End If
        Next KKI%
        KKJ% = KKJ% + 1
        CISL%(KKJ%) = CII%
        CANSL#(KKJ%) = CANTU#
      End If
20  Next U&
    '
    For KKI% = 1 To KKJ%
      If Abs(CANSL#(KKI%)) >= 0.05 Then
        CI% = CISL%(KKI%)
        LOTE$ = ""
        CMOV$ = CMO$
        If AMOVI$ = "MOVIREPA" Then CMOV$ = "*R" ' reparacion
        DOPRI$ = REMISEL$
        DOSEC$ = ""
        REFE$ = "Anulac.Remito " + REMISEL$
        VUNI# = 0
        MONE$ = ""
        NC% = CLI%
        CANTI = (-1) * CANSL#(KKI%)
        Call MOVISTO(FECHA%, CI%, LOTE$, CMOV$, DOPRI$, DOSEC$, REFE$, VUNI#, MONE$, NC%, DEPO%, CANTI)
      End If
    Next KKI%
    '
    NUREMIT& = Val(Mid$(REMISEL$, 4))
    FIN& = ULTREG&("REMPEN")
    J& = 0
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("REMPEN", U&)
      REMILE& = CVS(Mid$(X$, 3, 4))
      If REMILE& <> NUREMIT& Then
         J& = J& + 1
         Call GRAREG("REMPEN", X$, J&)
      End If
    Next U&
    Call SETULTREG("REMPEN", J&)
    '
50  NUREMIT& = Val(Mid$(REMISEL$, 4))
    FIN& = ULTREG&("CONSICLI")
    J& = 0
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("CONSICLI", U&)
      If Mid$(X$, 5, 3) = "RT." Then
        REMILE& = Val(Mid$(X$, 13, 8))
        If REMILE& = NUREMIT& Then
          Call REPLA(X$, String$(84, 0), 37, 84)
          Call REPLA(X$, "Anulacion Rto" + Str$(NUREMIT&), 63, 24)
          Call GRAREG("CONSICLI", X$, U&)
        End If
      End If
    Next U&
    '
    FIN& = ULTREG&("PEDDETA")
    J& = 0
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("PEDDETA", U&)
      If CVI(Mid$(X$, 7, 2)) = CLI% Then
        REMILE& = CVS(Mid$(X$, 105, 4))
        If REMILE& = NUREMIT& Then
          TOENTRE = CVS(Mid$(X$, 95, 4))
          SALPEN = CVS(Mid$(X$, 99, 4)) + TOENTRE
          If Asc(Mid$(X$, 27, 1)) < 9 Then
            Call REPLA(X$, Chr$(1), 27, 1)
          End If
          Call REPLA(X$, MKS$(0), 95, 4)
          Call REPLA(X$, MKS$(SALPEN), 99, 4)
          Call REPLA(X$, MKI$(0), 103, 2)
          Call REPLA(X$, MKS$(0), 105, 4)
          Call GRAREG("PEDDETA", X$, U&)
        End If
      End If
    Next U&
    '
    FIN& = ULTREG&("PEDDETP")
    J& = 0
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("PEDDETP", U&)
      If CVI(Mid$(X$, 7, 2)) = CLI% Then
        REMILE& = CVS(Mid$(X$, 105, 4))
        If REMILE& = NUREMIT& Then
          TOENTRE = CVS(Mid$(X$, 95, 4))
          SALPEN = CVS(Mid$(X$, 99, 4)) + TOENTRE
          If Asc(Mid$(X$, 27, 1)) < 9 Then
            Call REPLA(X$, Chr$(1), 27, 1)
          End If
          Call REPLA(X$, MKS$(0), 95, 4)
          Call REPLA(X$, MKS$(SALPEN), 99, 4)
          Call REPLA(X$, MKI$(0), 103, 2)
          Call REPLA(X$, MKS$(0), 105, 4)
          Call GRAREG("PEDDETP", X$, U&)
        End If
      End If
    Next U&
    '
99  Call CIERRARCH("*.*")
    Call BAJALDISCO
    Screen.MousePointer = 1
    Command25.Enabled = True
    '
End Sub

Private Sub Command3_Click()
    '
    Command3.Enabled = False
    '\\\OT-05-0827-WAR-21/11/05///
    ATX1% = ALTERNA%("Detalle de Entregas\Control de Facturación")
    Select Case ATX1%
      Case 1
        GoTo 5
      Case 2
        Call CONTROFAC
        GoTo 99
      Case Else
        GoTo 99
5   End Select
    '\\\OT-05-0827-WAR-FIN///
    ' detalle de entregas por cliente
    '
    Dim PUNMO$(32767)
    '
10  FEX = HOY(HO$)
    HAS% = FEX
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    DES% = FECHANUM(DESS$)
    VDEF$ = MKI$(DES%) + MKI$(HAS%)
    '
    OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
        GoTo 99
    End If
    '
    Call BLOQARCH("DETARTIC")
    '
    DES% = CVI(Left$(OBX$, 2))
    HAS% = CVI(Mid$(OBX$, 3, 2))
    If DES% > HAS% Then
        GoTo 10
    End If
    FEX = DES%: PERIO$ = FECHATEX$(FEX)
    FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    '
    Screen.MousePointer = 11
    Call HEADERS("DETCLIEN", "")
    Call HEADERS("DETCLIEN", "Periodo: " + PERIO$)
    '
    FIN& = ULTREG&("MOVISTO")
    J& = 0
    '
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = REGLEIDO$("MOVISTO", U&)
       FF% = CVI(Left$(X$, 2))
       If FF% >= DES% Then
          If FF% <= HAS% Then
             CLI% = CVI(Mid$(X$, 84, 2))
             If CLI% > 0 Then
                If CLI% <= 32767 Then
                   TIMO1$ = UCase$(Mid$(X$, 23, 2))
                   TIMO2$ = UCase$(Mid$(X$, 33, 2))
                   If TIMO1$ = "RT" Or TIMO2$ = "RT" Or TIMO1$ = "DV" Or TIMO2$ = "DV" Then
                      PUNMO$(CLI%) = PUNMO$(CLI%) + MKS$(U&)
                   End If
                End If
             End If
          End If
       End If
    Next U&
    '
    UREMO& = ULTREG&("MOVISTO")
    FIN& = ULTREG&("DEUDOR1")
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       Y$ = REGLEIDO$("DEUDOR1", U&)
       CLI% = CVI(Mid$(Y$, 1, 2))
       If CLI% > 0 Then
          If CLI% <= 32767 Then
             '
             For KKI% = 1 To Len(PUNMO$(CLI%)) Step 4
               REGI& = CVS(Mid$(PUNMO$(CLI%), KKI%, 4))
               If REGI& > 0 Then
                 If REGI& <= UREMO& Then
                   XI$ = REGLEIDO$("MOVISTO", REGI&)
                   CLI% = CVI(Mid$(XI$, 84, 2))
                   FFI% = CVI(Left$(XI$, 2))
                   REMIFAI$ = Mid$(XI$, 23, 10)
                   CANTUI# = CVD(Mid$(XI$, 96, 8)) - CVD(Mid$(X$, 88, 8))
                   For KKJ% = KKI% + 4 To Len(PUNMO$(CLI%)) Step 4
                     REGJ& = CVS(Mid$(PUNMO$(CLI%), KKJ%, 4))
                     If REGJ& > 0 Then
                       If REGJ& <= UREMO& Then
                         XJ$ = REGLEIDO$("MOVISTO", REGJ&)
                         If CLI% = CVI(Mid$(XJ$, 84, 2)) Then
                           If FFI% = CVI(Left$(XJ$, 2)) Then
                             If REMIFAI$ = Mid$(XJ$, 23, 10) Then
                               If Abs(CANTUI# - CVD(Mid$(XJ$, 88, 8)) + CVD(Mid$(XJ$, 96, 8))) < 0.005 Then
                                 Mid$(PUNMO$(CLI%), KKI%, 4) = MKS$(0)
                                 Mid$(PUNMO$(CLI%), KKJ%, 4) = MKS$(0)
                                 GoTo 19
                               End If
                             End If
                           End If
                         End If
                       End If
                     End If
                   Next KKJ%
                 End If
               End If
19           Next KKI%
             '
             For KKI% = 1 To Len(PUNMO$(CLI%)) Step 4
                REG& = CVS(Mid$(PUNMO$(CLI%), KKI%, 4))
                If REG& > 0 Then
                   If REG& <= UREMO& Then
                      X$ = REGLEIDO$("MOVISTO", REG&)
                      CLI% = CVI(Mid$(X$, 84, 2))
                      FF% = CVI(Left$(X$, 2))
                      CII% = CVI(Mid$(X$, 3, 2))
                      REMIFAC$ = Mid$(X$, 23, 10)
                      If Left$(REMIFAC$, 4) = "DV.0" Then
                        REMIFAC$ = Mid$(X$, 33, 10)
                      End If
                      CANTU# = CVD(Mid$(X$, 96, 8)) - CVD(Mid$(X$, 88, 8))
                      '
                      Z$ = REGBLAN$("DETCLIEN")
                      Call REPLA(Z$, MKI$(CLI%), 1, 2)
                      Call REPLA(Z$, MKI$(FF%), 3, 2)
                      Call REPLA(Z$, REMIFAC$, 5, 10)
                      Call REPLA(Z$, MKI$(CII%), 15, 2)
                      Call REPLA(Z$, MKD$(CANTU#), 17, 8)
                      J& = J& + 1
                      Call GRAREG("DETCLIEN", Z$, J&)
                   End If
                End If
             Next KKI%
             PUNMO$(CLI%) = ""
          End If
       End If
    Next U&
    '
    FIN& = 32767
    For U& = 1 To 32767
       Call TRACE(20, U&, FIN&)
       If PUNMO$(U&) <> "" Then
          For KKI% = 1 To Len(PUNMO$(U&)) Step 4
             REG& = CVS(Mid$(PUNMO$(U&), KKI%, 4))
             If REG& > 0 Then
                If REG& <= UREMO& Then
                   X$ = REGLEIDO$("MOVISTO", REG&)
                   CLI% = CVI(Mid$(X$, 84, 2))
                   FF% = CVI(Left$(X$, 2))
                   CII% = CVI(Mid$(X$, 3, 2))
                   REMIFAC$ = Mid$(X$, 23, 10)
                   CANTU# = CVD(Mid$(X$, 96, 8)) - CVD(Mid$(X$, 88, 8))
                   '
                   Z$ = REGBLAN$("DETCLIEN")
                   Call REPLA(Z$, MKI$(CLI%), 1, 2)
                   Call REPLA(Z$, MKI$(FF%), 3, 2)
                   Call REPLA(Z$, REMIFAC$, 5, 10)
                   Call REPLA(Z$, MKI$(CII%), 15, 2)
                   Call REPLA(Z$, MKD$(CANTU#), 17, 8)
                   J& = J& + 1
                   Call GRAREG("DETCLIEN", Z$, J&)
                End If
             End If
          Next KKI%
       End If
    Next U&
    '
    Call SETULTREG("DETCLIEN", J&)
    Call CIERRARCH("DETCLIEN")
    Call LIBARCH("DETCLIEN")
    Call FINAL("*DETACLI")
    '
99  Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    Command3.Enabled = True
    '
End Sub

Private Sub Command4_Click()
    '
    Command4.Enabled = False
    '
    ' detalle de entregas por artículo
    '
    Dim PUNMO$(32767)
    '
10  FEX = HOY(HO$)
    HAS% = FEX
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    DES% = FECHANUM(DESS$)
    VDEF$ = MKI$(DES%) + MKI$(HAS%)
    '
    OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
        GoTo 99
    End If
    '
    Call BLOQARCH("DETARTIC")
    '
    DES% = CVI(Left$(OBX$, 2))
    HAS% = CVI(Mid$(OBX$, 3, 2))
    If DES% > HAS% Then
        GoTo 10
    End If
    FEX = DES%: PERIO$ = FECHATEX$(FEX)
    FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    '
    Screen.MousePointer = 11
    Call HEADERS("DETARTIC", "")
    Call HEADERS("DETARTIC", "Periodo: " + PERIO$)
    '
    FIN& = ULTREG&("MOVISTO")
    J& = 0
    '
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = REGLEIDO$("MOVISTO", U&)
       FF% = CVI(Left$(X$, 2))
       If FF% >= DES% Then
         If FF% <= HAS% Then
           CLI% = CVI(Mid$(X$, 84, 2))
           If CLI% > 0 Then
             CI% = CVI(Mid$(X$, 3, 2))
             If CI% > 0 Then
               If CI% <= 32767 Then
                 TIMO1$ = UCase$(Mid$(X$, 23, 2))
                 TIMO2$ = UCase$(Mid$(X$, 33, 2))
                 If TIMO1$ = "RT" Or TIMO2$ = "RT" Or TIMO1$ = "DV" Or TIMO2$ = "DV" Then
                   PUNMO$(CI%) = PUNMO$(CI%) + MKS$(U&)
                 End If
               End If
             End If
           End If
         End If
       End If
    Next U&
    '
    UREMO& = ULTREG&("MOVISTO")
    FIN& = NUINV%
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       Y$ = REGLEIDO$("INVERT", U& + 1)
       CI% = CVI(Mid$(Y$, 17, 2))
       If CI% > 0 Then
          If CI% <= NUMAS% Then
             '
             For KKI% = 1 To Len(PUNMO$(CI%)) Step 4
               REGI& = CVS(Mid$(PUNMO$(CI%), KKI%, 4))
               If REGI& > 0 Then
                 If REGI& <= UREMO& Then
                   XI$ = REGLEIDO$("MOVISTO", REGI&)
                   CLI% = CVI(Mid$(XI$, 84, 2))
                   FFI% = CVI(Left$(XI$, 2))
                   REMIFAI$ = Mid$(XI$, 23, 10)
                   CANTUI# = CVD(Mid$(XI$, 96, 8)) - CVD(Mid$(X$, 88, 8))
                   For KKJ% = KKI% + 4 To Len(PUNMO$(CI%)) Step 4
                     REGJ& = CVS(Mid$(PUNMO$(CI%), KKJ%, 4))
                     If REGJ& > 0 Then
                       If REGJ& <= UREMO& Then
                         XJ$ = REGLEIDO$("MOVISTO", REGJ&)
                         If CLI% = CVI(Mid$(XJ$, 84, 2)) Then
                           If FFI% = CVI(Left$(XJ$, 2)) Then
                             If REMIFAI$ = Mid$(XJ$, 23, 10) Then
                               If Abs(CANTUI# - CVD(Mid$(XJ$, 88, 8)) + CVD(Mid$(XJ$, 96, 8))) < 0.005 Then
                                 Mid$(PUNMO$(CI%), KKI%, 4) = MKS$(0)
                                 Mid$(PUNMO$(CI%), KKJ%, 4) = MKS$(0)
                                 GoTo 19
                               End If
                             End If
                           End If
                         End If
                       End If
                     End If
                   Next KKJ%
                 End If
               End If
19           Next KKI%
             '
             For KKI% = 1 To Len(PUNMO$(CI%)) Step 4
                REG& = CVS(Mid$(PUNMO$(CI%), KKI%, 4))
                If REG& > 0 Then
                   If REG& <= UREMO& Then
                      X$ = REGLEIDO$("MOVISTO", REG&)
                      CLI% = CVI(Mid$(X$, 84, 2))
                      FF% = CVI(Left$(X$, 2))
                      CII% = CVI(Mid$(X$, 3, 2))
                      REMIFAC$ = Mid$(X$, 23, 10)
                      If Left$(REMIFAC$, 4) = "DV.0" Then
                        REMIFAC$ = Mid$(X$, 33, 10)
                      End If
                      CANTU# = CVD(Mid$(X$, 96, 8)) - CVD(Mid$(X$, 88, 8))
                      '
                      Z$ = REGBLAN$("DETARTIC")
                      Call REPLA(Z$, MKI$(CII%), 1, 2)
                      Call REPLA(Z$, MKI$(FF%), 3, 2)
                      Call REPLA(Z$, REMIFAC$, 5, 10)
                      Call REPLA(Z$, MKI$(CLI%), 15, 2)
                      Call REPLA(Z$, MKD$(CANTU#), 17, 8)
                      J& = J& + 1
                      Call GRAREG("DETARTIC", Z$, J&)
                   End If
                End If
             Next KKI%
          End If
       End If
    Next U&
    '
    Call SETULTREG("DETARTIC", J&)
    Call CIERRARCH("DETARTIC")
    Call LIBARCH("DETARTIC")
    Call FINAL("*DETARTI")
    '
99  Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    Command4.Enabled = True
    '
End Sub

Private Sub Command5_Click()
    '
    Command5.Enabled = False
10  FEX = HOY(HO$)
    HAS% = FEX
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    DES% = FECHANUM(DESS$)
    VDEF$ = MKI$(DES%) + MKI$(HAS%)
    '
    OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
        GoTo 99
    End If
    '
    DES% = CVI(Left$(OBX$, 2))
    HAS% = CVI(Mid$(OBX$, 3, 2))
    If DES% > HAS% Then
        GoTo 10
    End If
    FEX = DES%: PERIO$ = FECHATEX$(FEX)
    FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    '
    Screen.MousePointer = 11
    J& = 0
    FIN& = ULTREG&("MATVEN")
    '
    For U& = 1 To FIN&
        Call TRACE(20, U&, FIN&)
        X$ = REGLEIDO$("MATVEN", U&)
        FRI% = CVI(Mid$(X$, 7, 2))
        If FRI% >= DES% Then
            If FRI% <= HAS% Then
                NC% = CVI(Left$(X$, 2))
                DOPRI$ = "RT." + TRIM$(CVS(Mid$(X$, 23, 4)))
                CIXI% = CVI(Mid$(X$, 9, 2))
                CAN# = CVS(Mid$(X$, 19, 4))
                '
                J& = J& + 1
                Z$ = REGBLAN$("LIREMIT")
                Call REPLA(Z$, DOPRI$, 1, 10)
                Call REPLA(Z$, MKI$(FRI%), 11, 2)
                  RCCX$ = TRIM$(TEL2CLI$(NC%))
                  If RCCX$ = "" Then RCCX$ = RASOCLI$(NC%)
                Call REPLA(Z$, RCCX$, 13, 10)
                Call REPLA(Z$, MKI$(CIXI%), 23, 2)
                Call REPLA(Z$, MKD$(CAN#), 25, 8)
                Call GRAREG("LIREMIT", Z$, J&)
            End If
        End If
    Next U&
    '
    CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
    If Len(CMO$) <> 2 Then
       CMO$ = "RT"
    End If
    '
    FIN& = ULTREG&("MOVISTO")
    For U& = 1 To FIN&
        Call TRACE(20, U&, FIN&)
        X$ = REGLEIDO$("MOVISTO", U&)
        FRI% = CVI(Left$(X$, 2))
        If FRI% >= DES% Then
          If FRI% <= HAS% Then
             If UCase$(Mid$(X$, 21, 2)) = CMO$ Then
                NC% = CVI(Mid$(X$, 84, 2))
                DOPRI$ = Mid$(X$, 23, 10)
                CIXI% = CVI(Mid$(X$, 3, 2))
                CAN# = CVD(Mid$(X$, 96, 8)) - CVD(Mid$(X$, 88, 8))
                '
                J& = J& + 1
                Z$ = REGBLAN$("LIREMIT")
                Call REPLA(Z$, DOPRI$, 1, 10)
                Call REPLA(Z$, MKI$(FRI%), 11, 2)
                  RCCX$ = TRIM$(TEL2CLI$(NC%))
                  If RCCX$ = "" Then RCCX$ = RASOCLI$(NC%)
                Call REPLA(Z$, RCCX$, 13, 10)
                Call REPLA(Z$, MKI$(CIXI%), 23, 2)
                Call REPLA(Z$, MKD$(CAN#), 25, 8)
                Call GRAREG("LIREMIT", Z$, J&)
             End If
          End If
        End If
    Next U&
    '
    Call COPYSTRU("MOVISTO", "MOVIREPA")
    FIN& = ULTREG&("MOVIREPA")
    For U& = 1 To FIN&
        Call TRACE(20, U&, FIN&)
        X$ = REGLEIDO$("MOVIREPA", U&)
        FRI% = CVI(Left$(X$, 2))
        If FRI% >= DES% Then
          If FRI% <= HAS% Then
            NC% = CVI(Mid$(X$, 84, 2))
            DOPRI$ = Mid$(X$, 23, 10)
            CIXI% = CVI(Mid$(X$, 3, 2))
            CAN# = CVD(Mid$(X$, 96, 8)) - CVD(Mid$(X$, 88, 8))
            '
            J& = J& + 1
            Z$ = REGBLAN$("LIREMIT")
            Call REPLA(Z$, DOPRI$, 1, 10)
            Call REPLA(Z$, MKI$(FRI%), 11, 2)
              RCCX$ = TRIM$(TEL2CLI$(NC%))
              If RCCX$ = "" Then RCCX$ = RASOCLI$(NC%)
            Call REPLA(Z$, RCCX$, 13, 10)
            Call REPLA(Z$, MKI$(CIXI%), 23, 2)
            Call REPLA(Z$, MKD$(CAN#), 25, 8)
            Call GRAREG("LIREMIT", Z$, J&)
          End If
        End If
    Next U&
    '
    Call SETULTREG("LIREMIT", J&)
    Call CIERRARCH("LIREMIT")
    '
    Call HEADERS("LIREMIT", "")
    Call HEADERS("LIREMIT", "Periodo: " + PERIO$)
    '
    Screen.MousePointer = 1
    Call FINAL("*SOREMIT")
    '
99  Command5.Enabled = True
    '
End Sub

Private Sub Command6_Click()
    '
    Command6.Enabled = False
    Dim CISL%(2048), CANSL#(2048)
    '
    CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
    If Len(CMO$) <> 2 Then
       CMO$ = "RT"
    End If
    '
    Call CARSELMESA
    '
10  Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) = "" Then
       GoTo 99
    End If
    CORRESQ$ = MESANO$(MESASEL$, DES%, HAS%)
    '
    Call FRESHECHO("SELREMI")
    Screen.MousePointer = 11
    FIN& = ULTREG&("MOVISTO")
    J& = 0: REMIA$ = ""
    '
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = REGLEIDO$("MOVISTO", U&)
       FF% = CVI(Left$(X$, 2))
       If FF% >= DES% Then
          If FF% <= HAS% Then
             REMI$ = UCase$(Mid$(X$, 23, 10))
             If Val(Mid$(REMI$, 4)) > 0 Then
               If REMI$ <> REMIA$ Then
                 REMIA$ = REMI$
                 TIMO1$ = Mid$(X$, 21, 2)
                 If TIMO1$ = CMO$ Then
                   If UCase$(Mid$(X$, 33, 2)) <> "FC" Then 'excluye facturados
                     CLI% = CVI(Mid$(X$, 84, 2))
                     If CLI% > 0 Then
                       Z$ = Space$(40) + String$(8, 0)
                       Call REPLA(Z$, REMI$, 1, 10)
                       Call REPLA(Z$, RASOCLI$(CLI%), 11, 30)
                       Call REPLA(Z$, MKS$(U&), 41, 4)
                       Call REPLA(Z$, MKI$(FF%), 45, 2)
                       For KK& = 1 To J&
                         YZ$ = REGLEIDO$("SELREMI", KK&)
                         If Left$(YZ$, 40) = Left$(Z$, 40) Then GoTo 14
                       Next KK&
                       J& = J& + 1
                       Call GRAREG("SELREMI", Z$, J&)
                     End If
                   End If
                 End If
               End If
             End If
          End If
       End If
14  Next U&
    '
    FIN& = ULTREG&("MOVIREPA")
    REMIA$ = ""
    '
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = REGLEIDO$("MOVIREPA", U&)
       FF% = CVI(Left$(X$, 2))
       If FF% >= DES% Then
          If FF% <= HAS% Then
             REMI$ = UCase$(Mid$(X$, 23, 10))
             If Val(Mid$(REMI$, 4)) > 0 Then
               If REMI$ <> REMIA$ Then
                 REMIA$ = REMI$
                 TIMO1$ = Mid$(X$, 21, 2)
                 If TIMO1$ = "RP" Then
                   If UCase$(Mid$(X$, 33, 2)) <> "FC" Then 'excluye facturados
                     CLI% = CVI(Mid$(X$, 84, 2))
                     If CLI% > 0 Then
                       Z$ = Space$(40) + String$(8, 0)
                       Call REPLA(Z$, REMI$, 1, 10)
                       Call REPLA(Z$, RASOCLI$(CLI%), 11, 30)
                       Call REPLA(Z$, MKS$((-1) * U&), 41, 4)
                       Call REPLA(Z$, MKI$(FF%), 45, 2)
                       For KK& = 1 To J&
                         YZ$ = REGLEIDO$("SELREMI", KK&)
                         If Left$(YZ$, 40) = Left$(Z$, 40) Then GoTo 24
                       Next KK&
                       J& = J& + 1
                       Call GRAREG("SELREMI", Z$, J&)
                     End If
                   End If
                 End If
               End If
             End If
          End If
       End If
24  Next U&
    '
    Call SETULTREG("SELREMI", J&)
    Call CIERRARCH("SELREMI")
    Screen.MousePointer = 1
    '
    If J& < 1 Then
       Call MENSERR(24, "No se Encontraron Remitos Posibles de Anular\en el Período Elegido.")
       GoTo 10
    End If
    '
    Call SELECHO("SELREMI")
    REMISEL$ = VALACT1$("SELREMI")
    If TRIM$(REMISEL$) = "" Then
       GoTo 10
    End If
    '
    RERE$ = REGACT$("SELREMI")
    PRIREG& = CVS(Mid$(RERE$, 41, 4))
    CLIE$ = Mid$(RERE$, 11, 30)
    FEX = CVI(Mid$(RERE$, 45, 2))
    '
    If Abs(PRIREG&) < 1 Then GoTo 10
    '
    TXT$ = "Remito Número " + TRIM$(REMISEL$) + "\Cliente: " + TRIM$(CLIE$) + "\Fecha: " + FECHATEX$(FEX) + "\   \\Abandonar\Marcar 'Sin Cargo'"
    ALTE% = COMALTER%(TXT$)
    If ALTE% <> 2 Then GoTo 10
    '
    Call BLOQARCH("CONSICLI")
    '
    Screen.MousePointer = 11
    '
50  NUREMIT& = Val(Mid$(REMISEL$, 4))
    FIN& = ULTREG&("CONSICLI")
    J& = 0
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("CONSICLI", U&)
      If Mid$(X$, 5, 3) = "RT." Then
        REMILE& = Val(Mid$(X$, 13, 8))
        If REMILE& = NUREMIT& Then
          If Mid$(X$, 63, 9) = "Sin Cargo" Then
            Call COMUNI("Este Remito ya Estaba\Marcado como 'Sin Cargo'.")
            GoTo 99
          End If
          Call REPLA(X$, String$(84, 0), 37, 84)
          Call REPLA(X$, "Sin Cargo Rto" + Str$(NUREMIT&), 63, 24)
          Call GRAREG("CONSICLI", X$, U&)
        End If
      End If
    Next U&
    '
99  Call CIERRARCH("*.*")
    Call BAJALDISCO
    Screen.MousePointer = 1
    Command6.Enabled = True
    Exit Sub
    '
End Sub
'

Private Sub Command7_Click()
    Call HELPONLINE("ORPASTA")
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   MODO% = COMALTER%("Listado Material en Consignación\\Pantalla\Impresora\Total Pendiente\Cancelar")
   If MODO% < 1 Or MODO% > 3 Then GoTo 99
   '
   If MODO% = 3 Then GoTo 20     ' LISTADO general total pendiente
   '
10 Call SELECHO("DEUDOR1")
   NCX7% = Val(VALACT1$("DEUDOR1"))
   '
   If NCX7% > 0 Then
     If ECOARCH$("DEUDOR1", MKI$(NCX7%)) <> "" Then
       Call RECONCLI(NCX7%)
       JJ& = ULTREG&("RECONCLI")
       If JJ& < 1 Then
          Call MENSERR(24, "Sin Movimiento")
          GoTo 10
       End If
       '
       Call HEADERS("RECONCLI", "")
       Call HEADERS("RECONCLI", "Cliente: " + ECOARCH$("DEUDOR1", MKI$(NCX7%)))
       '
       If MODO% = 1 Then
           Call CONECRUN("*COCONCLI/" + TRIM$(Str$(NCX7%)), "Material en Consignacion (Clientes)/WAIT")
         Else
           Call CONECRUN("*LICONCLI", "Material en Consignacion (Clientes)")
       End If
       '
       GoTo 10
       '
     End If
   End If
   GoTo 99
   '
20 Screen.MousePointer = 11
   Dim ACLIX%(32767)
   For U& = 1 To ULTREG&("CONSICLI")
      X$ = REGLEIDO$("CONSICLI", U&)
      NCLIE% = CVI(Left$(X$, 2))
      If NCLIE% > 0 Then
        If NCLIE% <= 32767 Then
          ACLIX%(NCLIE%) = 1
        End If
      End If
   Next U&
   '
   KK& = 0
   Call COPYSTRU("RECONCLI", "REPEFAC")
   For NCLIE% = 1 To 32760
     If ACLIX%(NCLIE%) > 0 Then
        Call RECONCLI(NCLIE%)
        For U& = 1 To ULTREG&("RECONCLI")
          X$ = REGLEIDO$("RECONCLI", U&)
          KK& = KK& + 1
          Call GRAREG("REPEFAC", X$, KK&)
        Next U&
     End If
   Next NCLIE%
   Call SETULTREG("REPEFAC", KK&)
   Call CIERRARCH("REPEFAC")
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   '
   Call FINAL("*REPEFAC")
   '
99 Command8.Enabled = True
End Sub

Private Sub Command9_Click()
   Command9.Enabled = False
10 Call SELECHO("DEUDOR1")
   NCX7% = Val(VALACT1$("DEUDOR1"))
   '
   If NCX7% > 0 Then
     If ECOARCH$("DEUDOR1", MKI$(NCX7%)) <> "" Then
       Call RECONCLI(NCX7%)
       JJ& = ULTREG&("RECONCLI")
       If JJ& < 1 Then
          Call MENSERR(24, "Sin Movimiento")
          GoTo 10
       End If
       '
       Call SELECHO("RECONMAT")
       CODMAT$ = VALACT1$("RECONMAT")
       CI% = CODINT%(CODMAT$)
       CADEVU = Val(VALACT2$("RECONMAT"))
       '
       If TRIM$(CODMAT$) <> "" Then
         TTXX$ = "Devolución Código " + TRIM$(CODMAT$) + "\Cantidad = " + TRIM$(VALACT2$("RECONMAT")) + " Unidades."
         If COMALTER%(TTXX$ + "\\Confirmar\Cancelar") = 1 Then
            HO% = HOY(HOS$)
            HP = 100 * Val(Left$(Time$, 2)) + Val(Mid$(Time$, 4, 2))
            USER% = USNBR% Mod 100
            CTRXX$ = MKI$(HO%) + MKI$(Int(HP)) + Chr$(USER%) + String$(2, 0)
            '
            RECONSI$ = REGBLAN$("CONSICLI")
            Call REPLA(RECONSI$, String$(12, 0), 109, 12)
            Call REPLA(RECONSI$, MKI$(NCX7%), 1, 2)
            Call REPLA(RECONSI$, MKI$(HO%), 3, 2)
            Call REPLA(RECONSI$, "Devolucion", 5, 18)
            Call REPLA(RECONSI$, String$(12, 0), 109, 12)
            Call REPLA(RECONSI$, CTRXX$, 122, 7)
            Call REPLA(RECONSI$, "Devolucion", 23, 12)
            Call REPLA(RECONSI$, MKI$(CI%), 35, 2)
            Call REPLA(RECONSI$, MKD$(CDbl(Abs(CADEVU))), 45, 8)
            Call REPLA(RECONSI$, MEDIDAS$, 109, 12)
            Call REGAPP("CONSICLI", RECONSI$)
            Call CIERRARCH("CONSICLI")
            '
            Call MOVISTO(HO%, CI%, "", "RT", "Dev-Cons", "Dev-Cons", RASOCLI$(NCX7%), VUNI#, "$", NCX7%, 1, CADEVU)
            Call CIERRARCH("MOVISTO")
            '
         End If
       End If
       '
       Call FRESHECHO("RECONMAT")
       GoTo 10
       '
     End If
   End If
   '
   Command9.Enabled = True
End Sub
'
Sub RECONCLI(NCX7%)
   '
   Dim SALCONSI(32767)
   Call SETULTREG("RECONCLI", 0)
   Call SETULTREG("RECONMAT", 0)
   '
   If NCX7% > 0 Then
     If ECOARCH$("DEUDOR1", MKI$(NCX7%)) <> "" Then
       Screen.MousePointer = 11
       FIN& = ULTREG&("CONSICLI")
       For KLK& = 1 To 32767: SALCONSI(KLK&) = 0: Next KLK&
       For U& = 1 To FIN&
         Call TRACE(20, U&, FIN&)
         XX$ = REGLEIDO$("CONSICLI", U&)
         NCX8% = CVI(Left$(XX$, 2))
         If NCX8% = NCX7% Then
           CIXI% = CVI(Mid$(XX$, 35, 2))
           If CIXI% > 0 Then
             If CIXI% <= NUMAS% Then
               CACON = CVD(Mid$(XX$, 37, 8)) - CVD(Mid$(XX$, 45, 8))
               SALCONSI(CIXI%) = SALCONSI(CIXI%) + CACON
             End If
           End If
         End If
       Next U&
       '
       JJ& = 0
       FIN& = NUINV%
       For U& = 1 To NUINV%
         Call TRACE(20, U&, FIN&)
         II$ = REGLEIDO$("INVERT", U& + 1)
         CIXI% = CVI(Mid$(II$, 17, 2))
         If CIXI% > 0 Then
           If CIXI% <= NUMAS% Then
             If SALCONSI(CIXI%) > 0 Then
               XX$ = REGBLAN$("RECONCLI")
               YY$ = REGBLAN$("RECONMAT")
               Call REPLA(XX$, MKI$(NCX7%), 1, 2)
               Call REPLA(XX$, MKI$(CIXI%), 3, 2)
               Call REPLA(XX$, MKS$(SALCONSI(CIXI%)), 5, 4)
               Call REPLA(XX$, String$(16, "."), 9, 16)
               Call REPLA(YY$, MKI$(CIXI%), 1, 2)
               Call REPLA(YY$, MKS$(SALCONSI(CIXI%)), 3, 4)
               JJ& = JJ& + 1
               Call GRAREG("RECONCLI", XX$, JJ&)
               Call GRAREG("RECONMAT", YY$, JJ&)
             End If
           End If
         End If
       Next U&
       Call SETULTREG("RECONCLI", JJ&)
       Call CIERRARCH("RECONCLI")
       Call SETULTREG("RECONMAT", JJ&)
       Call CIERRARCH("RECONMAT")
       Screen.MousePointer = 1
       '
     End If
   End If

End Sub

Private Sub Form_Activate()
    '
    Command19.Enabled = False
    If CONTROL$("REMSTAN", "REPLAORI") = "SI" Then
      Command19.Enabled = True
    End If
    '
    If EXISTE%("COCUMENT.EXE") > 0 Then
      Command11.Enabled = True
    End If
    '
End Sub

