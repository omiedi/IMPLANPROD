VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form MENUACRE07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C7D9CC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestión Compras - Ctas. Ctes."
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   4995
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   4995
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   17
      Top             =   1320
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0050765A&
      ForeColor       =   &H80000008&
      Height          =   4425
      Left            =   4095
      ScaleHeight     =   4395
      ScaleWidth      =   945
      TabIndex        =   12
      Top             =   0
      Width           =   975
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H000C93DC&
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   19
         Top             =   3480
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0076C9F5&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -120
            TabIndex        =   20
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command8 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   140
         Picture         =   "MENUACRE07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Listados de Documentos Clasificados"
         Top             =   2205
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command29 
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   140
         Picture         =   "MENUACRE07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Ayuda Flexoft en Línea"
         Top             =   900
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   140
         Picture         =   "MENUACRE07.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   210
         Width           =   650
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   105
         TabIndex        =   15
         Top             =   1785
         Visible         =   0   'False
         Width           =   645
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C7D9CC&
      Caption         =   "GESTIÓN DE COMPRAS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3765
      Left            =   105
      TabIndex        =   2
      Top             =   105
      Width           =   3900
      Begin VB.Frame Frame8 
         BackColor       =   &H00C7D9CC&
         Caption         =   "CONTROL DE GESTION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   210
         TabIndex        =   8
         Top             =   2310
         Width           =   3480
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   420
            ScaleHeight     =   615
            ScaleWidth      =   2610
            TabIndex        =   9
            Top             =   420
            Width           =   2640
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
               Height          =   630
               Left            =   0
               Picture         =   "MENUACRE07.frx":075E
               Style           =   1  'Graphical
               TabIndex        =   10
               ToolTipText     =   "Control General de Gestión de Compras"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Control General de Gestión de Compras"
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
               Left            =   700
               TabIndex        =   11
               Top             =   100
               Width           =   1905
            End
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C7D9CC&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1905
         Left            =   240
         TabIndex        =   3
         Top             =   315
         Width           =   3480
         Begin VB.PictureBox Picture8 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   735
            ScaleHeight     =   615
            ScaleWidth      =   2295
            TabIndex        =   4
            Top             =   945
            Width           =   2325
            Begin VB.CommandButton Command14 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   630
               Left            =   0
               Picture         =   "MENUACRE07.frx":0A68
               Style           =   1  'Graphical
               TabIndex        =   1
               ToolTipText     =   "Emitir Ordenes de Pago"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "Orden de Pago en Cta.Corriente"
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
               Left            =   700
               TabIndex        =   5
               Top             =   100
               Width           =   1485
            End
         End
         Begin VB.PictureBox Picture14 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   420
            ScaleHeight     =   615
            ScaleWidth      =   2295
            TabIndex        =   6
            Top             =   420
            Width           =   2325
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
               Height          =   630
               Left            =   0
               Picture         =   "MENUACRE07.frx":0BB2
               Style           =   1  'Graphical
               TabIndex        =   0
               ToolTipText     =   "Registrar Facturas de Compras"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "Facturación de Proveedores"
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
               Left            =   700
               TabIndex        =   7
               Top             =   50
               Width           =   1590
            End
         End
         Begin VB.Label LabelFuenteAnchoFijo 
            Caption         =   "Para aplicar fuente ancho fijo"
            Height          =   255
            Left            =   600
            TabIndex        =   21
            Top             =   1560
            Visible         =   0   'False
            Width           =   2295
         End
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   750
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "MENUACRE07.frx":0FF4
      TabIndex        =   18
      Top             =   -120
      Visible         =   0   'False
      Width           =   1590
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3600
      OleObjectBlob   =   "MENUACRE07.frx":106E
      Top             =   1080
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   5880
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu op_comer 
      Caption         =   "Transacciones"
      Begin VB.Menu fac_provee 
         Caption         =   "Factura de Proveedores"
      End
      Begin VB.Menu orpago 
         Caption         =   "Ordenes de Pago"
      End
      Begin VB.Menu plh1 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu cons_inf 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu res_porcta 
         Caption         =   "Control General de Gestión de Compras"
      End
      Begin VB.Menu mnuCtaCte 
         Caption         =   "Resumen de Cuenta Corriente"
      End
   End
End
Attribute VB_Name = "MENUACRE07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Function ValorDeCompraSegunCantDetfacom(FACTU$, NPROV, MON%)
   'RECORRO LA ORDEN DE COMPRA
   'TOMO LA CANTIDAD DEL DETFACOM DEL ARTICULO Y VEO SI TIENE O.C.
   'MULTIPLICO = IMPORTE DE COMPRA X CANTIDAD DE DETFACOM
   'VERIFICO MONEDA Y EN TODO CASO CAMBIO A MONEDA DE DETFACOM O LA DE CAJABANC
   VALORCOMPRATOTAL = 0
   SQLX$ = "SELECT SUM(CANTFACTURADA) AS CANTI, COD_INTE, NUME_OCOM FROM  DetfacomSQL WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE FACTURAPROV = '" & FACTU$ & "'"
   SQLX$ = SQLX$ + " AND NUME_CLI = " & NPROV
   SQLX$ = SQLX$ + " GROUP BY COD_INTE,NUME_OCOM"
   SQLX$ = SQLX$ + " ORDER BY COD_INTE"
   Set RS = READSET(SQLX$)
   With RS
     Do While Not .EOF
       CI1% = XVALO(!Cod_Inte)
       NOC& = XVALO(!NUME_OCOM)
       CANTI = XVALO(!CANTI)
       Condi$ = " Cod_Inte = " & CI1%
       Condi$ = Condi$ + " AND Nume_Ocom = " & NOC&
       Condi$ = Condi$ + " AND Npro_Ocom = " & NPROV
       PRECIOCOMPRA = XVALO(VALOBADA("OCOMDETA", "PRE_UNIT", Condi$, "NOMESS"))
       VALORCOMPRATOTAL = VALORCOMPRATOTAL + (PRECIOCOMPRA * CANTI)
       .MoveNext
     Loop
   End With
   Condi$ = "  Nume_Ocom = " & NOC&
   Condi$ = Condi$ + " AND Npro_Ocom = " & NPROV

   MONECOM% = XVALO(VALOBADA("OCOMENCA", "MONE_EMIS", Condi$, "NOMESS"))
   COEFCON = 1
   If MONECOM% <> MON% Then
      COEFCON = TICAMONE(MONECOM%) / TICAMONE(MON%)
   End If
   VALORCOMPRATOTAL = VALORCOMPRATOTAL * COEFCON
   ValorDeCompraSegunCantDetfacom = VALORCOMPRATOTAL
End Function

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    OPP% = COMALTER%("Gestión de Acreedores Comerciales\\Minimizar\Cerrar\Cancelar")
    If OPP% = 1 Then
        WindowState = 1
      ElseIf OPP% = 2 Then
        Call FINAL("")
        Exit Sub
    End If
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   If DERACCE%("ORPAGO", "Orden de Pago en Cta. Corriente") > 1 Then
'     Call RECLAPAG

     ORPAGO07.Show 1
   End If
   Command14.Enabled = True
End Sub


Private Sub Command21_Click()
    Command21.Enabled = False
    IMPREFAC.Show 1
    Command21.Enabled = True
End Sub




Private Sub Command22_Click()
    '
    Command22.Enabled = False
    If DERACCE%("FACPROV", "Facturación de Proveedores") > 1 Then
      FORFAC07.Show 1
    End If
    Command22.Enabled = True
    '
End Sub

Private Sub Command29_Click()
   Call HELPONLINE("ACRE_GES")
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   If DERACCE%("COGESCO", "Control General de Gestión de Compras") > 1 Then
     'Call COGECOM07.Command2_Click
     COGECOM07.Show 1
   End If
   Command3.Enabled = True
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   OPXX% = COMALTER%("Reporte Documentos Clasificados:\\Compras y Pagos\Asientos Contables")
   If OPXX% = 1 Then
        Call FINAL("?COMPAGCLASI") ' Compras y Pagos Clasificados
     ElseIf OPXX% = 2 Then
        Call FINAL("?ASICONCLASI") ' Asientos Contables Clasificados
   End If
   Command8.Enabled = True
End Sub

Private Sub cta_porpagar_Click()
   Call COGECOM07.Command3_Click
End Sub

Private Sub ctipo_fac_Click()
   Call FRESHALL
   Call CARSELMESA
   Call SELECHO("SELMESA")
   MESASEL$ = VALACT1$("SELMESA")
   If TRIM$(MESASEL$) <> "" Then
     REMECO07.MESAN = MESASEL$
     Call REMECO07.Command3_Click
   End If
End Sub

Private Sub ctipo_ncre_Click()
   Call FRESHALL
   Call CARSELMESA
   Call SELECHO("SELMESA")
   MESASEL$ = VALACT1$("SELMESA")
   If TRIM$(MESASEL$) <> "" Then
     REMECO07.MESAN = MESASEL$
     Call REMECO07.Command5_Click
   End If
End Sub

Private Sub ctipo_ndeb_Click()
   Call FRESHALL
   Call CARSELMESA
   Call SELECHO("SELMESA")
   MESASEL$ = VALACT1$("SELMESA")
   If TRIM$(MESASEL$) <> "" Then
     REMECO07.MESAN = MESASEL$
     Call REMECO07.Command4_Click
   End If
End Sub

Private Sub ctipo_opago_Click()
   Call FRESHALL
   Call CARSELMESA
   Call SELECHO("SELMESA")
   MESASEL$ = VALACT1$("SELMESA")
   If TRIM$(MESASEL$) <> "" Then
     REMECO07.MESAN = MESASEL$
     Call REMECO07.Command2_Click
   End If
End Sub

Private Sub exit_Click()
   Call Command1_Click
End Sub

Private Sub fac_provee_Click()
   Call Command22_Click
End Sub

Private Sub Form_Activate()
   Call refresca_botones(Me)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     Openforms = DoEvents
   Wend
End Sub

Private Sub Form_Load()
   '
   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then Call FINAL("")
   '
   EJER% = 1
   '

   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If EPAC$ <> "" Then
      Caption = Caption + "  -  " + EPAC$
   End If
   Call GRATITFOR(Caption)
        If ControlMenu%(Me.Name, Me) < 1 Then Call FINAL("")

'   UTILIZA_SKIN% = 1
'   Call APLICACIONSKINS(Me, Picture1)
'   Call InstalaFuenteAnchoFijo
   '
   Call ABRECONEXION
   '
   If InStr(UCase$(EPAC$), "FERVI") > 0 Then
      HOII% = HOY(HOS$)
      If HOII% <= FECHANUM("28/11/06") Then
         If TRIM$(Control$("MIGRAFER", "CHECARTE")) = "" Then
            Show
            Call RECUCHECAR
            Call FINAL("")
         End If
      End If
   End If
   '
'   'VALIDACION POR SI NO PAGA
'   If InStr(EPAC$, "DAVTECNICA") > 0 Then
'        FF% = FECHANUM("30/06/20")
'        If FF% <= HOY(HO$) Then
'          Call Form_Load
'        End If
'   End If
   
10 CUMA$ = Control$("AMAPRO", "ECUECON")
   If CUMA$ = "" Or CUEINT%(CUMA$) < 1 Then
     Call MENSERR(24, "Falta Cuenta Contable por Omisión\Correspondiente a 'Acreedores Comerciales'`.\  \Establezca la Cuenta Antes de Continuar.")
     Call CIERRARCH("*.*")
     Call CONECRUN("ARCHIG07/OPMASPRO", "Archivos Maestros")
     Call FINAL("")
   End If
   '
   'ORDEN DE PAGO A PROVEEDORES
   '
   If Left$(APLINVO$, 6) = "ORPAGO" Then                'ORDEN DE PAGO
       Command14_Click
       'ORPAGO07.Show 1
       Call FINAL("")
     ElseIf Left$(APLINVO$, 9) = "OPCIORPAG" Then       'OPCIONES VARIAS - ORDENES DE PAGOS
       Call ORPAGO07.Command6_Click
       Call FINAL("")
     ElseIf Left$(APLINVO$, 11) = "CONGENORPAG" Then    'CONFIGURACION GENERAL - ORDEN DE PAGO
       Call ORPAGO07.Command7_Click
       Call FINAL("")
     ElseIf Left$(APLINVO$, 8) = "ANUORPAG" Then        'ANULACION DE ORDENES DE PAGOS
       Call OPORPA07.Command25_Click
       Call FINAL("")
     ElseIf Left$(APLINVO$, 9) = "VERORPAG" Then        'VER Y REIMPRIMIR ORDEN DE PAGO
       Call OPORPA07.Command23_Click
       Call FINAL("")
     ElseIf Left$(APLINVO$, 10) = "IMPCREACRE" Then     'ANULACION DE ORDENES DE PAGOS
       Call OPORPA07.Command3_Click
       Call FINAL("")
     ElseIf Left$(APLINVO$, 9) = "VISIMPACR" Then       'VISUALIZAR IMPUTACION DE CREDITOS
       Call OPORPA07.Command11_Click
       Call FINAL("")
     ElseIf Left$(APLINVO$, 10) = "AJUSALACRE" Then     'AJUSTE SALDO ACREEDORES
       Call OPORPA07.Command4_Click
       Call FINAL("")
     ElseIf Left$(APLINVO$, 9) = "COMCLIPRO" Then       'COMPENSACION CLIENTE PROVEEDOR
       Call OPORPA07.Command10_Click
       Call FINAL("")
     ElseIf Left$(APLINVO$, 10) = "TRANSORPAG" Then     'TRANSFERENCIA DE ORDEN DE PAGO
       Call OPORPA07.Command13_Click
       Call FINAL("")
     ElseIf Left$(APLINVO$, 10) = "DESAPORPAG" Then     'DESAPLICAR ORDEN DE PAGO
       Call OPORPA07.Command12_Click
       Call FINAL("")
     ElseIf Left$(APLINVO$, 10) = "CONLISOPAG" Then     'CONSULTAS Y LISTADOS
       Call ORPAGO07.Command6_Click
       Call FINAL("")
     ElseIf Left$(APLINVO$, 10) = "LISORPAG" Then       'LISTADO DE ORDENES DE PAGO
       Call OPORPA07.Command5_Click
       Call FINAL("")
     '
     'FACTURACION DE PROVEEDORES
     '
     ElseIf Left$(APLINVO$, 7) = "FACTURA" Then         'FACTURACION DE PROVEEDORES
       FORFAC07.Show 1
       Call FINAL("")
     ElseIf Left$(APLINVO$, 10) = "OPCIFACPRO" Then     'OPCIONES VARIAS FACTURA DE PROVEDORES
       FACOCOMP07.Show 1
       Call FINAL("")
     ElseIf Left$(APLINVO$, 8) = "ANUFACOM" Then        'ANULACION DE COMPROBANTES
       Call FACOCOMP07.Command25_Click
       Call FINAL("")
     ElseIf Left$(APLINVO$, 10) = "IMPUCREPRO" Then     'IMPUTACION DE CREDITOS ACREEDORES
       Call FACOCOMP07.Command3_Click
       Call FINAL("")
     ElseIf Left$(APLINVO$, 10) = "AJUSALACRE" Then     'AJUSTE DE SALDOS ACREEDORES
       Call FACOCOMP07.Command4_Click
       Call FINAL("")
     ElseIf Left$(APLINVO$, 9) = "CONFACPRO" Then       'CONFIGURACION GENERAL - FACTURA DE PROVEEDORES
       Call FORFAC07.Command8_Click
       Call FINAL("")
     '
     'CONTROL DE GESTION COMERCIAL DE COMPRAS
     '
     ElseIf Left$(APLINVO$, 9) = "CGCOMPRAS" Then         'CONTROL DE GESTION DE COMPRAS
       COGECOM07.Show 1
       Call FINAL("")
     ElseIf Left$(APLINVO$, 9) = "CONESTCOM" Then         'CONTROL ESTADO
       COGECOM07.Show 1
       Call FINAL("")
     ElseIf Left$(APLINVO$, 8) = "CUEPAGAR" Then          'CUENTAS POR PAGAR
       Call COGECOM07.Command3_Click
       Call FINAL("")
     ElseIf Left$(APLINVO$, 10) = "SALDEUPROV" Then       'SALDO Y DEUDAS ACREEDORES
       Call COGECOM07.Command5_Click
       Call FINAL("")
     ElseIf Left$(APLINVO$, 10) = "DETCUEPROV" Then       'DETALLE CUENTA CORRIENTE PROVEEDOR
       Call COGECOM07.Command7_Click
       Call FINAL("")
     ElseIf Left$(APLINVO$, 9) = "SUBESTCOM" Then         'RESUMENES MENSUALES - SUBDIARIOS Y ESTADISTICAS
       Call COGECOM07.Command2_Click
       Call FINAL("")
   End If
   '
   UTILIZA_SKIN% = 1
   Call APLICACIONSKINS(Me, Picture1)
   Call InstalaFuenteAnchoFijo
   '
   Call CREARCAMPOS_DETFACOMSQL
   If InStr(EPAC$, "AHP") > 0 Then
     Call CREARCAMPOS_FACTURAYGASTOS
   End If
   '
   If NIVCONFI% >= 5 Then
      Command8.Visible = True
   End If
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
   Call BAJALDISCO
   Call FINAL("")
End Sub

Private Sub help_Click()
   'CON
'   If CLACONTRA% > 0 Then
'        'ESTE CODIGO ES PARA DAR DE ALTA UNOS SALDOS QUE PASO DEL PROVEEDOR, A MODO DE VALOR DE COMPRA
'        IMPOGASTO = 0
'        IMPOFAC = 854.8
'        FACTURA_ASOCIADA$ = ""
'20      RNDX$ = RANDSTRING$(-10)
'        FACTU$ = "FC-E-0001-00007605"
'        CONDI1$ = "FACTURAPROV ='" & RNDX$ & "'"
'        If CantRegistros("FACTURAYGASTOS", CONDI1$, "NOMESS") > 0 Then GoTo 20
'        Call CreaRegistro("FACTURAYGASTOS", "FACTURAPROV", RNDX$, "NOMESS")
'        '
'        Call ACTUREGISTRO("FACTURAYGASTOS", "VAL_COMPRA", CONDI1$, 2862.21, RAFE&, "NOMESS")
'        Call ACTUREGISTRO("FACTURAYGASTOS", "FECHEMISI", CONDI1$, FETEXCOR1$(FECHANUM("30/06/14")), RAFE&, "NOMESS")
'        Call ACTUREGISTRO("FACTURAYGASTOS", "IMPORTE_FACTURA", CONDI1$, FORMATNUM$(IMPOFAC, "F16.5"), RAFE&, "NOMESS")
'        Call ACTUREGISTRO("FACTURAYGASTOS", "IMPO_GASTO", CONDI1$, FORMATNUM$(IMPOGASTO, "F16.5"), RAFE&, "NOMESS")
'        Call ACTUREGISTRO("FACTURAYGASTOS", "FACTURAPROV_ASOC", CONDI1$, FACTURA_ASOCIADA$, RAFE&, "NOMESS")
'
'        Call ACTUREGISTRO("FACTURAYGASTOS", "TIPO_FACTURA", CONDI1$, "M", RAFE&, "NOMESS")
'        Call ACTUREGISTRO("FACTURAYGASTOS", " Nume_Cli", CONDI1$, 10014, RAFE&, "NOMESS")
'        Call ACTUREGISTRO("FACTURAYGASTOS", "MoneVal", CONDI1$, 2, RAFE&, "NOMESS")
'        Call ACTUREGISTRO("FACTURAYGASTOS", "TipoCam", CONDI1$, 9.083, RAFE&, "NOMESS")
'        Call ACTUREGISTRO("FACTURAYGASTOS", "Nume_Cli_Asoc", CONDI1$, 0, RAFE&, "NOMESS")
'        '
'        Call ACTUREGISTRO("FACTURAYGASTOS", "FACTURAPROV", CONDI1$, FACTU$, RAFE&, "NOMESS")
'   '--------------------------------------2
'        IMPOGASTO = 0
'        IMPOFAC = 7600.46
'        FACTURA_ASOCIADA$ = ""
'30      RNDX$ = RANDSTRING$(-10)
'        FACTU$ = "OC-E-0001-72826100"
'        CONDI1$ = "FACTURAPROV ='" & RNDX$ & "'"
'        If CantRegistros("FACTURAYGASTOS", CONDI1$, "NOMESS") > 0 Then GoTo 30
'        Call CreaRegistro("FACTURAYGASTOS", "FACTURAPROV", RNDX$, "NOMESS")
'        '
'        Call ACTUREGISTRO("FACTURAYGASTOS", "VAL_COMPRA", CONDI1$, (-717.49), RAFE&, "NOMESS")
'        Call ACTUREGISTRO("FACTURAYGASTOS", "FECHEMISI", CONDI1$, FETEXCOR1$(FECHANUM("30/06/14")), RAFE&, "NOMESS")
'        Call ACTUREGISTRO("FACTURAYGASTOS", "IMPORTE_FACTURA", CONDI1$, FORMATNUM$(IMPOFAC, "F16.5"), RAFE&, "NOMESS")
'        Call ACTUREGISTRO("FACTURAYGASTOS", "IMPO_GASTO", CONDI1$, FORMATNUM$(IMPOGASTO, "F16.5"), RAFE&, "NOMESS")
'        Call ACTUREGISTRO("FACTURAYGASTOS", "FACTURAPROV_ASOC", CONDI1$, FACTURA_ASOCIADA$, RAFE&, "NOMESS")
'
'        Call ACTUREGISTRO("FACTURAYGASTOS", "TIPO_FACTURA", CONDI1$, "M", RAFE&, "NOMESS")
'        Call ACTUREGISTRO("FACTURAYGASTOS", " Nume_Cli", CONDI1$, 10030, RAFE&, "NOMESS")
'        Call ACTUREGISTRO("FACTURAYGASTOS", "MoneVal", CONDI1$, 2, RAFE&, "NOMESS")
'        Call ACTUREGISTRO("FACTURAYGASTOS", "TipoCam", CONDI1$, 9.083, RAFE&, "NOMESS")
'        Call ACTUREGISTRO("FACTURAYGASTOS", "Nume_Cli_Asoc", CONDI1$, 0, RAFE&, "NOMESS")
'        '
'        Call ACTUREGISTRO("FACTURAYGASTOS", "FACTURAPROV", CONDI1$, FACTU$, RAFE&, "NOMESS")
'
'   '--------------------------------------3
'        IMPOGASTO = 0
'        IMPOFAC = 8944.21
'        FACTURA_ASOCIADA$ = ""
'40      RNDX$ = RANDSTRING$(-10)
'        FACTU$ = "FC-E-0001-00134020"
'        CONDI1$ = "FACTURAPROV ='" & RNDX$ & "'"
'        If CantRegistros("FACTURAYGASTOS", CONDI1$, "NOMESS") > 0 Then GoTo 40
'        Call CreaRegistro("FACTURAYGASTOS", "FACTURAPROV", RNDX$, "NOMESS")
'        '
'        Call ACTUREGISTRO("FACTURAYGASTOS", "VAL_COMPRA", CONDI1$, 5449.68, RAFE&, "NOMESS")
'        Call ACTUREGISTRO("FACTURAYGASTOS", "FECHEMISI", CONDI1$, FETEXCOR1$(FECHANUM("30/06/14")), RAFE&, "NOMESS")
'        Call ACTUREGISTRO("FACTURAYGASTOS", "IMPORTE_FACTURA", CONDI1$, FORMATNUM$(IMPOFAC, "F16.5"), RAFE&, "NOMESS")
'        Call ACTUREGISTRO("FACTURAYGASTOS", "IMPO_GASTO", CONDI1$, FORMATNUM$(IMPOGASTO, "F16.5"), RAFE&, "NOMESS")
'        Call ACTUREGISTRO("FACTURAYGASTOS", "FACTURAPROV_ASOC", CONDI1$, FACTURA_ASOCIADA$, RAFE&, "NOMESS")
'
'        Call ACTUREGISTRO("FACTURAYGASTOS", "TIPO_FACTURA", CONDI1$, "M", RAFE&, "NOMESS")
'        Call ACTUREGISTRO("FACTURAYGASTOS", " Nume_Cli", CONDI1$, 10008, RAFE&, "NOMESS")
'        Call ACTUREGISTRO("FACTURAYGASTOS", "MoneVal", CONDI1$, 2, RAFE&, "NOMESS")
'        Call ACTUREGISTRO("FACTURAYGASTOS", "TipoCam", CONDI1$, 9.069, RAFE&, "NOMESS")
'        Call ACTUREGISTRO("FACTURAYGASTOS", "Nume_Cli_Asoc", CONDI1$, 0, RAFE&, "NOMESS")
'        '
'        Call ACTUREGISTRO("FACTURAYGASTOS", "FACTURAPROV", CONDI1$, FACTU$, RAFE&, "NOMESS")
'
'    '--------------------------------------4
'        IMPOGASTO = 0
'        IMPOFAC = 2863.54
'        FACTURA_ASOCIADA$ = ""
'50      RNDX$ = RANDSTRING$(-10)
'        FACTU$ = "FC-E-0001-01373385"
'        CONDI1$ = "FACTURAPROV ='" & RNDX$ & "'"
'        If CantRegistros("FACTURAYGASTOS", CONDI1$, "NOMESS") > 0 Then GoTo 50
'        Call CreaRegistro("FACTURAYGASTOS", "FACTURAPROV", RNDX$, "NOMESS")
'        '
'        Call ACTUREGISTRO("FACTURAYGASTOS", "VAL_COMPRA", CONDI1$, (-2823.29), RAFE&, "NOMESS")
'        Call ACTUREGISTRO("FACTURAYGASTOS", "FECHEMISI", CONDI1$, FETEXCOR1$(FECHANUM("30/06/14")), RAFE&, "NOMESS")
'        Call ACTUREGISTRO("FACTURAYGASTOS", "IMPORTE_FACTURA", CONDI1$, FORMATNUM$(IMPOFAC, "F16.5"), RAFE&, "NOMESS")
'        Call ACTUREGISTRO("FACTURAYGASTOS", "IMPO_GASTO", CONDI1$, FORMATNUM$(IMPOGASTO, "F16.5"), RAFE&, "NOMESS")
'        Call ACTUREGISTRO("FACTURAYGASTOS", "FACTURAPROV_ASOC", CONDI1$, FACTURA_ASOCIADA$, RAFE&, "NOMESS")
'
'        Call ACTUREGISTRO("FACTURAYGASTOS", "TIPO_FACTURA", CONDI1$, "M", RAFE&, "NOMESS")
'        Call ACTUREGISTRO("FACTURAYGASTOS", " Nume_Cli", CONDI1$, 10007, RAFE&, "NOMESS")
'        Call ACTUREGISTRO("FACTURAYGASTOS", "MoneVal", CONDI1$, 2, RAFE&, "NOMESS")
'        Call ACTUREGISTRO("FACTURAYGASTOS", "TipoCam", CONDI1$, 9.069, RAFE&, "NOMESS")
'        Call ACTUREGISTRO("FACTURAYGASTOS", "Nume_Cli_Asoc", CONDI1$, 0, RAFE&, "NOMESS")
'        '
'        Call ACTUREGISTRO("FACTURAYGASTOS", "FACTURAPROV", CONDI1$, FACTU$, RAFE&, "NOMESS")
'     'PROCESO DE RECUPERAR FACTURA DE MATERIALES EN BASE A LA FACTURAS DE GASTOS
'      '
'      'RECORRER LA TABLA DE FACTURA Y GASTOS.
'      'VALIDAR SI LA FACTURA DE PROVEDOR ASOCIADO NO ESTA COMO FACTURA DE PROVEEDOR COMO TIPO FACTURA M
'      'SI NO ESTA GENERAR UN NUEVO REGISTO
'      'EL VALOR DE LA ORDEN DE COMPRA VERIFICAR EL TEMA DE LA MONEDA SI DIFIERE CON LA ORDEN DE COMPRA
'
''      'RECORRER LA TABLA DE FACTURA Y GASTOS.
''      SQLX$ = "SELECT * FROM FACTURAYGASTOS WITH(NOLOCK) WHERE TIPO_FACTURA = 'G' ORDER BY FACTURAPROV_ASOC"
''      Set RS = READSET(SQLX$)
''      With RS
''         Do While Not .EOF
''           FACTURAPROVASOCIADO$ = SAFETEXT$(!FACTURAPROV_ASOC)
''           CONDI$ = "FACTURAPROV = '" & FACTURAPROVASOCIADO$ & "' AND TIPO_FACTURA = 'M'"
''           If CantRegistros&("FACTURAYGASTOS", CONDI$, "NOMESS") < 1 Then
''            SQLX$ = "SELECT * FROM FACTURAYGASTOS WHERE NUME_CLI < 1 "
''            Dim RS1 As ADODB.Recordset
''            Set RS1 = New ADODB.Recordset
''25          On Error Resume Next
''            RS1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
''            If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
''               On Error GoTo 0
''               Call CapturaErrorOpen
''               GoTo 25
''            End If
''            On Error GoTo 0
''            With RS1
''              .AddNew
''              !FECHEMISI = RS!FECHEMISI
''              FACTURA$ = SAFETEXT$(RS!FACTURAPROV_ASOC)
''              CONDI$ = "TIPOMOVIM = 'COMPRA' AND CODICOM_LAR = '" & FACTURA$ & "' AND NUPROVE = " & XVALO(RS!Nume_Cli_Asoc)  ' VALOR NETO"
''              !IMPORTE_FACTURA = XVALO(VALOBADA("CAJABANC", "SUM(IDEBECOMP)", CONDI$, "NOMESS"))
''                 MON% = XVALO(VALOBADA("CAJABANC", "MONE_EMIS", CONDI$, "NOMESS"))
''              !MoneVal = XVALO(VALOBADA("CAJABANC", "MONE_EMIS", CONDI$, "NOMESS"))
''              !TipoCam = XVALO(VALOBADA("CAJABANC", "TIPO_CAM", CONDI$, "NOMESS"))
''              !IMPO_GASTO = 0
''              !FACTURAPROV_ASOC = ""
''              !TIPO_FACTURA = "M"
''              !Nume_Cli = RS!Nume_Cli_Asoc
''              CONDI$ = "FACTURAPROV =  '" & FACTURA$ & "' AND NUME_CLI = " & XVALO(RS!Nume_Cli_Asoc)
''              !NUMEBORE = XVALO(VALOBADA("DETFACOMSQL", "NUMEBORE", CONDI$, "NOMESS"))
''              !FACTURAPROV = FACTURA$
''              !VAL_COMPRA = ValorDeCompraSegunCantDetfacom(FACTURA$, XVALO(RS!Nume_Cli_Asoc), MON%)
''              !Nume_Cli_Asoc = 0
''              OPX% = COMALTER%("Factura: " & FACTURA$ & " Orden de Compra: " & NUMEBORE & " \\Cancelar \Continuar")
''              If OPX% < 2 Then GoTo 30
''
''              .Update
''            End With
''           End If
''
''           .MoveNext
''         Loop
''      End With
''30    RS.Close
'' On Error Resume Next
'' RS1.Close: Set RS1 = Nothing
'' On Error GoTo 0
'   End If
   Call Command29_Click
End Sub

Private Sub mnuCtaCte_Click()
    Call ORPAGO07.Command4_Click
End Sub

Private Sub orpago_Click()
   Call Command14_Click
End Sub

Private Sub print1_Click()
   Call RECUEP07.Command3_Click
End Sub

Private Sub print2_Click()
   COGECOM07.Option2.Value = True
   Call COGECOM07.Command5_Click
End Sub

Private Sub res_porcta_Click()
  Call Command3_Click
End Sub

Private Sub scr1_Click()
   Call COGECOM07.Command7_Click
End Sub

Private Sub scr2_Click()
   COGECOM07.Option1.Value = True
   Call COGECOM07.Command5_Click
End Sub

Private Sub sub_compras_Click()
   Call FRESHALL
   Call CARSELMESA
   Call SELECHO("SELMESA")
   MESASEL$ = VALACT1$("SELMESA")
   If TRIM$(MESASEL$) <> "" Then
     REMECO07.MESAN = MESASEL$
     Call REMECO07.Command16_Click
   End If
End Sub

Private Sub sub_ivacom_Click()
   Call FRESHALL
   Call CARSELMESA
   Call SELECHO("SELMESA")
   MESASEL$ = VALACT1$("SELMESA")
   If TRIM$(MESASEL$) <> "" Then
     REMECO07.MESAN = MESASEL$
     Call REMECO07.Command21_Click
   End If
End Sub

Sub RECUCHECAR()
    '
    Call COMUNI("FLEXOFT Debe Controlar y Quizá Recuperar\la Cartera de Cheques y Valores.\   \El Proceso Puede Demandar Varios Minutos.\  \No Utilizar las Aplicaciones\de Cuentas Corrientes Acreedoras hasta Finalizar.\  \Pulse 'Aceptar Mensaje' para Comenzar.")
    '
100 'RECUPERACION DE MOVIMIENTOS ANULADOS
    Call APERBASDAT("CABAN")
'   With CAJJJABANC
'      On Error Resume Next
'      .MoveLast
'      If Err Then
'         On Error GoTo 0
'         GoTo 200
'      End If
'      FIN& = .RecordCount 'NO FUNCIONA
'      JJ& = 0
'      .MoveFirst
'      Do While Not .EOF
'         JJ& = JJ& + 1
'         Call TRACE(20, JJ&, FIN&)
'         If Mid$(!codicom_lar, 6, 4) = "anul" Then
'            If Abs(!ValAbsComp) > 0 Then
'               '.Edit
'               !ValAbsComp = 0
'               !IDEBECOMP = 0
'               !IHABECOMP = 0
'               !BASEIMPO = 0
'               .Update
'            End If
'         End If
'      .MoveNext
'      Loop
'   End With
'   CAJJJABANC.Close
'   Set CAJJJABANC = Nothing
    SQLX$ = "Update CAJABANC SET ValAbsComp = 0"
    SQLX$ = SQLX$ + " ,IDEBECOMP = 0"
    SQLX$ = SQLX$ + " ,IHABECOMP = 0"
    SQLX$ = SQLX$ + " ,BASEIMPO = 0"
    SQLX$ = SQLX$ + " WHERE SUBSTRING(codicom_lar, 6, 4) = 'anul'"
    SQLX$ = SQLX$ + " AND Abs(ValAbsComp) > 0 "
    On Error Resume Next
     FLEXCONN.Execute (FILTSQL$(SQLX$))
    If Err Then
      On Error GoTo 0
      GoTo 200
    End If
    On Error GoTo 0
   '
'200 ' ESTO ES PARA RECUPERAR MOVIMIENTOS 'AS' EN FERVI
   '
   Call APERBASDAT("CABAN")
'   SQLX$ = "SELECT * FROM CAJABANC "
'   SQLX$ = SQLX$ + "WHERE NuProve > 0 "
'   SQLX$ = SQLX$ + "AND CodiCom_Lar like 'AS-%' "
   'Set Cabasel1 = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
   
   '
'   With Cabasel1
'     On Error Resume Next
'130    .MoveFirst
'       If Err = 0 Then
'         .Delete
'         GoTo 130
'       End If
'     On Error GoTo 0
'   End With

200 ' ESTO ES PARA RECUPERAR MOVIMIENTOS 'AS' EN FERVI
   SQLX$ = "DELETE FROM CAJABANC "
   SQLX$ = SQLX$ + "WHERE NuProve > 0 "
   SQLX$ = SQLX$ + "AND CodiCom_Lar like 'AS-%' "
    FLEXCONN.Execute (FILTSQL$(SQLX$))
   '
'   SQLX$ = "SELECT * FROM SACREPEN "
'   SQLX$ = SQLX$ + "WHERE NuProve > 0 "
'   SQLX$ = SQLX$ + "AND CodiCom_Lar like 'AS-%' OR CodiCom_Lar like 'FF-%' "
'   Set Cabasel1 = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
'   '
'   With Cabasel1
'     On Error Resume Next
'140    .MoveFirst
'       If Err = 0 Then
'         .Delete
'         GoTo 140
'       End If
'     On Error GoTo 0
'   End With
   '
   SQLX$ = "DELETE FROM SACREPEN "
   SQLX$ = SQLX$ + "WHERE NuProve > 0 "
   SQLX$ = SQLX$ + " AND CodiCom_Lar like 'AS-%'"
   SQLX$ = SQLX$ + " OR CodiCom_Lar like 'FF-%' "
    FLEXCONN.Execute (FILTSQL$(SQLX$))
   '
   Call GRAFACRE("AS")
   '
300 ' RECUPERACION CARTERA DE CHEQUES
   Dim IDECH$()
   HOII% = HOY(HOS$)
   Call APERBASDAT("CABAN")
   Screen.MousePointer = 11
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + "WHERE TipoMovim = 'CHRECI'"
   SQLX$ = SQLX$ + " AND DescriMov = '" & IDEVA$ & "'"
   'Set CHETEMP1 = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
   Dim CHETEMP1 As ADODB.Recordset
   Set CHETEMP1 = New ADODB.Recordset
25 On Error Resume Next
   CHETEMP1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
            '
   FIN& = ULTREG&("BADACHE")
   ReDim IDECH$(FIN&)
   For U& = 1 To FIN&
      Label1 = U&
      Label1.Visible = True
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("BADACHE", U&)
      IDEVA$ = TRIM$(Left$(XX$, 40))
         IDECH$(U&) = IDEVA$
      IMPORT# = CVD(Mid$(XX$, 49, 8))
      FEMICHE% = CVI(Mid$(XX$, 57, 2))
      FEVECHE% = CVI(Mid$(XX$, 59, 2))
      frecep% = CVI(Mid$(XX$, 61, 2))
      DORECEP$ = TRIM$(Mid$(XX$, 63, 10))
      DESTI$ = TRIM$(Mid$(XX$, 103, 30))
      FENTREGA% = CVI(Mid$(XX$, 133, 2))
      OB1$ = TRIM$(Mid$(XX$, 135, 40))
      OB2$ = TRIM$(Mid$(XX$, 175, 40))
      OB3$ = TRIM$(Mid$(XX$, 215, 40))
      '
      ' BUSCA POR IDEVA$
      With CHETEMP1
          On Error GoTo 0
'         .FindFirst "DescriMov = '" & IDEVA$ & "' " LA LLEVE ARRIBA A LA CONSULTA
'         If .NoMatch = False Then
          .MoveFirst
          If Err < 1 Then
              On Error GoTo 0
              Call REPLA(XX$, "**", 255, 2)
              Call GRAREG("BADACHE", XX$, U&)
              GoTo 19    ' LO ENCONTRO
            Else
              On Error GoTo 0
              GoTo 29
         End If
         '
19       If FENTREGA% > 0 Then
            ' MARCAR LA SALIDA
'            .Edit
            !Status = 2
            !FEDEPOSENT = CVDAT(FENTREGA%)
            !FECHENTRE = CVDAT(FENTREGA%)
            If TRIM$(" " & !Document) = "" Then !Document = "Mig.Datos"
            If TRIM$(" " & !DESTICHEQ) = "" Then !DESTICHEQ = DESTI$
            If TRIM$(" " & !REFESALCHE) = "" Then !REFESALCHE = OB1$
            If TRIM$(" " & !IDENSALCHE) = "" Then !IDENSALCHE = OB2$ ' AJUSTI$(HOS$ + " - " + Left$(Time$, 5) + " hs - Term: " + NOMTER$, 48)
            If TRIM$(" " & !USUASALCHE) = "" Then !USUASALCHE = OB3$ ' AJUSTI$("Usuario: " + TRIM$(USERNAME$), 48)
            .Update
         End If
      End With
      '
29 Next U&
   CHETEMP1.Close
   Set CHETEMP1 = Nothing

   '
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + "WHERE TipoMovim = 'CHRECI'"
   SQLX$ = SQLX$ + "AND Status = 1"
'   Set CHETEMP1 = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
   '
   Set CHETEMP1 = New ADODB.Recordset
28 On Error Resume Next
   CHETEMP1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 28
   End If
   On Error GoTo 0

   With CHETEMP1
      On Error Resume Next
      .MoveLast
      FIN& = .RecordCount
      If Err Then
         On Error GoTo 0
         GoTo 40
      End If
      On Error GoTo 0
      '
      .MoveFirst
      JJ& = 0
      FIN1& = UBound(IDECH$)
      Do While Not .EOF
         JJ& = JJ& + 1
         Call TRACE(20, JJ&, FIN&)
         IDEVA$ = !DESCRIMOV
         For U& = 1 To FIN1&
           If IDECH$(U&) = IDEVA$ Then
              XX$ = REGLEIDO$("BADACHE", U&)
              IMPORT# = CVD(Mid$(XX$, 49, 8))
              FEMICHE% = CVI(Mid$(XX$, 57, 2))
              FEVECHE% = CVI(Mid$(XX$, 59, 2))
              frecep% = CVI(Mid$(XX$, 61, 2))
              DORECEP$ = TRIM$(Mid$(XX$, 63, 10))
              DESTI$ = TRIM$(Mid$(XX$, 103, 30))
              FENTREGA% = CVI(Mid$(XX$, 133, 2))
              OB1$ = TRIM$(Mid$(XX$, 135, 40))
              OB2$ = TRIM$(Mid$(XX$, 175, 40))
              OB3$ = TRIM$(Mid$(XX$, 215, 40))
              '
              If FENTREGA% < 1 Then
'                   .Edit
                   !Status = 1
                   !FEDEPOSENT = 0
                   !FECHENTRE = 0
                   !Document = ""
                   !DESTICHEQ = DESTI$
                   !REFESALCHE = OB1$
                   !IDENSALCHE = OB2$ ' AJUSTI$(HOS$ + " - " + Left$(Time$, 5) + " hs - Term: " + NOMTER$, 48)
                   !USUASALCHE = OB3$ ' AJUSTI$("Usuario: " + TRIM$(USERNAME$), 48)
                   .Update
                 Else
'                   .Edit
                   !Status = 2
                   !FEDEPOSENT = CVDAT(FENTREGA%)
                   !FECHENTRE = CVDAT(FENTREGA%)
                   If TRIM$(" " & !Document) = "" Then !Document = "Mig.Datos"
                   If TRIM$(" " & !DESTICHEQ) = "" Then !DESTICHEQ = DESTI$
                   If TRIM$(" " & !REFESALCHE) = "" Then !REFESALCHE = OB1$
                   If TRIM$(" " & !IDENSALCHE) = "" Then !IDENSALCHE = OB2$ ' AJUSTI$(HOS$ + " - " + Left$(Time$, 5) + " hs - Term: " + NOMTER$, 48)
                   If TRIM$(" " & !USUASALCHE) = "" Then !USUASALCHE = OB3$ ' AJUSTI$("Usuario: " + TRIM$(USERNAME$), 48)
                   .Update
              End If
              GoTo 39
           End If
         Next U&
         '
         ' NO LO ENCONTRO EN BADACHE
'         .Edit
         !Status = 2
         !FEDEPOSENT = !fechconta
         !FECHENTRE = !fechconta
         If TRIM$(" " & !Document) = "" Then !Document = "Mig.Datos"
         If TRIM$(" " & !DESTICHEQ) = "" Then !DESTICHEQ = "Nunca Entró"
         If TRIM$(" " & !REFESALCHE) = "" Then !REFESALCHE = "No Figurado en BADACHE"
         If TRIM$(" " & !IDENSALCHE) = "" Then !IDENSALCHE = AJUSTI$(HOS$ + " - " + Left$(Time$, 5) + " hs - Term: " + NOMTER$, 48)
         If TRIM$(" " & !USUASALCHE) = "" Then !USUASALCHE = AJUSTI$("Usuario: " + TRIM$(USERNAME$), 48)
         .Update
39    .MoveNext
      Loop
      '
   End With
   CHETEMP1.Close
   Set CHETEMP1 = Nothing
   '
40 Screen.MousePointer = 1
   HOII% = HOY(HOS$)
   Call GRACONTROL$("MIGRAFER", "CHECARTE", HOS$)
   Call COMUNI("Fin de Control y Recuperación")
   '
   
End Sub

Sub GRAFACRE(Optional TIMOFIL$)
   '
   Static YAINDEXCHE%
   '
   Dim FEVEN%(10)
   Dim VENCIPESOS#(10)
   '
   Dim CAMP%, CUEMP%(2048), MEPA%(2048)
   For i& = 1 To ULTREG&("MEPAGO")
       F0$ = REGLEIDO$("MEPAGO", i&)
       CAMP% = i&
       MEPA%(CAMP%) = CVI(Left$(F0$, 2))
       CUEX$ = Mid$(F0$, 45, 12)
       CUEMP%(CAMP%) = CUEINT%(CUEX$)
       If TRIM$(CODCUE$(CUEMP%(CAMP%))) <> TRIM$(CUEX$) Then
          Call MENSERR(24, "Cuenta Contable Incorrecta\para Medio de Pago '" + TRIM$(Mid$(F0$, 3, 12)) + "'\-\Resultados Imprevisibles.")
          Call FINAL("")
       End If
   Next i&
   ' ESTA CONSULTA SE LLEVO ABAJO Y SE UNIFICO(POR QUE HABIA UN FINDFIRST)

'   Call APERBASDAT("CABAN")
'   Screen.MousePointer = 11
'   SQLX$ = "SELECT * FROM CAJABANC "
'   SQLX$ = SQLX$ + "WHERE TipoMovim = 'CHRECI' "
'   SQLX$ = SQLX$ + "AND Status = 1 "
'   SQLX$ = SQLX$ + "AND OPCIMOVIM = 'CH' "
'   SQLX$ = SQLX$ + " AND DescriMov = '" & IDEVA$ & "'"
'   SQLX$ = SQLX$ + " AND ABS(ValAbsComp - " & IMCHEX$ & ") < 0.005 "
   'Set CHETEMP1 = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
              
              '
   NOCLOSE% = 0
   Call APERBASDAT("CABAN")
   NOCLOSE% = 1
   FIN& = ULTREG&("CUECORR")
   For U& = 1 To FIN&
      Label1 = U&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("CUECORR", U&)
      FF% = CVI(Left$(XX$, 2))
      NC% = CVI(Mid$(XX$, 5, 2))
      NROX$ = Mid$(XX$, 7, 18)
      NRECIB$ = Left$(NROX$, 3) + Mid$(NROX$, 9, 1) + "-" + Mid$(NROX$, 14)
      TIDOCOM$ = Left$(NROX$, 2)
      '
      ' ESTO ES PARA RECUPERAR AJUSTES DE SALDOS
      If TIMOFIL$ <> "" Then
         If TIMOFIL$ = "AS" And TIDOCOM$ = "FF" Then GoTo 12
         If TIDOCOM$ <> TIMOFIL$ Then
            GoTo 15
         End If
      End If
      '
12    SUC% = 0
      NRO& = XVALO(Mid$(NROX$, 11, 8))
      Observa$ = Mid$(XX$, 25, 24)
      ICONTAX$ = Mid$(XX$, 199, 80)
      If TIDOCOM$ = "OP" Or TIDOCOM$ = "FF" Then
          GoSub 30
        Else
          GoSub 20
      End If
      '
15 Next U&
   GoSub GRAPLIP
Exit Sub
   '
   '
20 NC1% = (-1) * NC%
   '
   CoComCor$ = NRECIB$
   CoComLar$ = NROX$
'   CodiEmp% = 0
'   CodiSuc% = SUC%
   NumeClie = 0
   NumeProv% = NC%
   FechEmi = CVDAT(FF%)
   '
   Referen$ = rasocli$(NC1%)
   FeContab = CVDAT(FF%)
   FeAcredi = CVDAT(FF%)
   NumeTal = 0
   NumeCom = NRO&
      If Observa$ = "" Then Observa$ = rasocli$(NC1%)
   NumCai$ = ""
   FechCai = CVDAT(FECHANUM("01/01/80"))
   InclIva$ = ""
   TOFA0# = 0
   '
   ' graba detalle por imputación de compras
   For KKI% = 1 To Len(ICONTAX$) Step 10
     CuentaCo% = CVI(Mid$(ICONTAX$, KKI%, 2))
     CentroCos$ = ""
     ItConcep$ = "" 'P/ conceptos
     ImpoDeb# = Abs(CVD(Mid$(ICONTAX$, KKI% + 2, 8)))
     If TIDOCOM$ = "AS" Then
        If KKI% = 1 Then
            ImpoDeb# = CVD(Mid$(XX$, 123, 8)) - CVD(Mid$(XX$, 115, 8))
          Else
            ImpoDeb# = 0
        End If
     End If
       If TIDOCOM$ = "NC" Or TIDOCOM$ = "DP" Then ImpoDeb# = (-1) * ImpoDeb#
     '
     If CuentaCo% > 0 Then
       If Abs(ImpoDeb#) >= 0.005 Then
         DEMOVX$ = DENOCUE$(CuentaCo%)
         TiMovi$ = "COMPRA"
         OpMovi$ = TIDOCOM$
         VaMovi$ = Mid$(NROX$, 4, 1)
         Descrip$ = TRIM$(DEMOVX$) + " - " + rasocli$(NC1%)
         TOFA0# = TOFA0# + ImpoDeb#
         Call GRACABAN
       End If
     End If
     CentroCos$ = ""
   Next KKI%
   '
   ' grabar importes de IVA
   SUMIVA# = 0
   IEXEN# = ROUND(CVD(Mid$(XX$, 59, 8)))
   If XVALO(IEXEN#) >= 0.005 Then   ' GRABAR IMPORTE EXENTO
     TiMovi$ = "IVA"
     OpMovi$ = "CrFi"
     VaMovi$ = "Exento"
     ImpoDeb# = 0
     BasImpo# = ROUND(IEXEN#)
     Descrip$ = "IVA B.Imp. Exento " + NRECIB$
     CuentaCo% = CUECOINT%(Control$("", "CUEIVA-0"))
     TOFA0# = TOFA0# + ImpoDeb#
     Call GRACABAN
   End If
   '
   INOGRA# = ROUND(CVD(Mid$(XX$, 67, 8)))
   If XVALO(INOGRA#) >= 0.005 Then   ' GRABAR IMPORTE NO GRAVADO
     TiMovi$ = "IVA"
     OpMovi$ = "CrFi"
     VaMovi$ = "No Grav"
     ImpoDeb# = 0
     BasImpo# = INOGRA#
     Descrip$ = "IVA B.Imp. No Grav." + NRECIB$
     CuentaCo% = CUECOINT%(Control$("", "CUEIVA-0"))
     TOFA0# = TOFA0# + ImpoDeb#
     Call GRACABAN
   End If
   '
'  HASTA AQUI LLEGUE *************************************
   For KKI% = 75 To 114 Step 8
     TiMovi$ = "IVA"
     OpMovi$ = "CrFi"
     VaMovi$ = ""
     Descrip$ = "IVA Credito Fiscal " + NRECIB$
     ImpoDeb# = Abs(CVD(Mid$(XX$, KKI%, 8)))
     If TIDOCOM$ = "AS" Then
        ImpoDeb# = ImpoDeb# * Sgn(CVD(Mid$(XX$, 123, 8)) - CVD(Mid$(XX$, 115, 8)))
     End If
     If Abs(ImpoDeb#) < 0.005 Then GoTo 29
     '
     BasImpo# = 0
     If KKI% = 75 Then
         VaMovi$ = "21.00 %"
         CuentaCo% = CUECOINT%(Control$("", "CUEIVA-0"))
       ElseIf KKI% = 83 Then
         VaMovi$ = "Aduana"
         CuentaCo% = CUECOINT%(Control$("", "CUEIVA-4"))
       ElseIf KKI% = 91 Then
         VaMovi$ = "Per-IVA"
         Descrip$ = "Percepcion I.V.A. " + NRECIB$
         CuentaCo% = CUECOINT%(Control$("", "CUEIVA-5"))
       ElseIf KKI% = 99 Then
         TiMovi$ = "GAN"
         VaMovi$ = "Per-Gan"
         Descrip$ = "Percepcion Ganancias " + NRECIB$
         CuentaCo% = CUECOINT%(Control$("", "CUEIVA-6"))
       ElseIf KKI% = 107 Then
         TiMovi$ = "I-BRU"
         VaMovi$ = "PERC"
         Descrip$ = "Percepcion I.Brutos " + NRECIB$
' ojo que esto está mal
         CuentaCo% = CUECOINT%(Control$("", "CUEIVA-6"))
     End If
     '
     If TIDOCOM$ = "NC" Or TIDOCOM$ = "DP" Then ImpoDeb# = (-1) * ImpoDeb#
     If Abs(ImpoDeb#) >= 0.005 Then
       TOFA0# = TOFA0# + ImpoDeb#
       Call GRACABAN
     End If
     BasImpo# = 0
29 Next KKI%
   '
   ' graba factura en c.corriente proveedor
   CuentaCo% = CUECOINT%(CUEMADRE$(-1 * NC%))
   ImpoDeb# = (-1) * TOFA0#
     If TIDOCOM$ = "NC" Or TIDOCOM$ = "DP" Then ImpoDeb# = (-1) * ImpoDeb#
   TiMovi = "FCOM"
   OpMovi = Left$(Text11, 2)
   VaMovi = Label20
   Descrip$ = CoComLar$ + " - " + rasocli$(NC1%)
   FeAcredi = CVDAT(FF%)
   Call GRACABAN
   '
   ' Grabar Vencimientos de Facturas ******************************************
   SIGNO% = 1
   If TIDOCOM$ = "NC" Or TIDOCOM$ = "DP" Then SIGNO% = (-1)
     'IBruTot# = Abs(TOFAC#(0))
   IBruTot# = TOFA0#
   IIvaTot# = SUMIVA#
   '
   PRICONVEN% = -1
   SALBRUTO# = TOFA0# * SIGNO%
   TXVENCI$ = Mid$(XX$, 139, 60)
   For KKI% = 1 To 6
      FEVEN%(KKI%) = CVI(Mid$(TXVENCI$, 10 * KKI% - 9, 2))
      VENCIPESOS#(KKI%) = CVD(Mid$(TXVENCI$, 10 * KKI% - 7, 8))
   Next KKI%
   '
   For KKI% = 1 To 6
      If FEVEN(KKI%) <> 0 Then
        If Abs(VENCIPESOS#(KKI%)) >= 0.005 Then
           NuorVen% = KKI%
           FeVenc = CVDAT(FEVEN(KKI%))
           ImVenc# = VENCIPESOS#(KKI%)
           Call GRACREPEN
         End If
      End If
   Next KKI%
   '
Return
'
30 ' GRABA ORDENES DE PAGO

    CoComCor$ = NRECIB$
    CoComLar$ = NROX$
    IMPOP# = ROUND(CVD(Mid$(XX$, 59, 8)))
'    CodiEmp% = 0
'    CodiSuc% = SUC%
    NumeClie = 0
    NumeProv% = NC%
    FechEmi = CVDAT(FF%)
    Referen$ = rasocli$((-1) * NC%)
    FeContab = CVDAT(FF%)
    FeAcredi = CVDAT(FF%)
    CuentaCo% = CUECOCLI%(NC1%)
    TiMovi$ = "OPAG"
    OpMovi$ = "OP"
    VaMovi$ = "CCORR"
    Descrip$ = "Orden de Pago " + NRECIB$
    NumeTal = COBRACT%
    NumeCom = NRO&
    StatInit = 0
    ImpoDeb# = IMPOP#
    BasImpo# = IMPOP#    ' aqui va entonces el neto sin IVA
    Observa$ = ""
    If Observa$ = "" Then Observa$ = rasocli$(NC1%)
    '
    If TICUEPRO%((-1) * NC%) = 0 Then
         Call GRACABAN
       Else
         RFG$ = RECFILT$("DETPAGAS", 3, MKS$(NRO&))
         For K2& = 1 To Len(RFG$) Step 4
           K1& = CVS(Mid$(RFG$, K2&, 4))
           Z$ = REGLEIDO$("DETPAGAS", K1&)
           If CVI(Mid$(Z$, 7, 2)) = FF% Then
             CuentaCo% = CVI(Mid$(Z$, 17, 2))
             ImpoDeb# = CVD(Mid$(Z$, 9, 8))
             VaMovi$ = "CGAST"
             Referen$ = DENOCUE$(CuentaCo%)
             CentroCos$ = ""
             ItConcep$ = "" 'P/ conceptos
             Call GRACABAN
           End If
         Next K2&
    End If
    '
    'grabar cheques entregados
    If YAINDEXCHE% = 0 Then
        XX1$ = REGLEIDO$("CHEQUEMI", 1)
        Call GRAREG("CHEQUEMI", XX1$, 1)
        Call CIERRARCH("CHEQUEMI")
        YAINDEXCHE% = 1
    End If
    '
    RFGH$ = RECFILT$("CHEQUEMI", 7, MKS$(NRO&))
    For KKUU& = 1 To Len(RFGH$) Step 4
       RECHE& = CVS(Mid$(RFGH$, KKUU&, 4))
       Z$ = REGLEIDO$("CHEQUEMI", RECHE&)
       If CVI(Mid$(Z$, 1, 2)) = NC% Then     ' COINCIDE PROVEEDOR
          BANCO$ = AJUSTI$(Mid$(Z$, 15, 16), 18)
          NUCHE$ = AJUSTI$(Mid$(Z$, 3, 12), 14)
            FEX = CVI(Mid$(Z$, 49, 2))
            FEVEN1% = FEX
          FEVECH$ = FECHATEX$(FEX)
          IDEVA$ = BANCO$ + NUCHE$ + FEVECH$
          IMCHE# = CVD(Mid$(Z$, 33, 8))
          FECHE% = CVI(Mid$(Z$, 31, 2))
          '
          'BUSCAR SI ESTÁ INGRESADO
          ' CASO AFIRMATIVO:
          ' 1) MARCAR LA SALIDA DE CARTERA
          ' 2) PONER CUENTA = CUENTA DE CARTERA DE VALORES
          VUELTA% = 0: ENCUEN% = 0
          IMCHEX$ = TRIM$(FORMATNUM$(IMCHE#, "F14.2"))
          If IMCHEX$ = "" Then IMCHEX$ = "0"
          
'          SQLX$ = "DescriMov = '" & IDEVA$ & "' "
'          SQLX$ = SQLX$ + "AND ABS(ValAbsComp - " & IMCHEX$ & ") < 0.005 "
'          SQLX$ = SQLX$ + "AND Status = 1"

          'AQUI SE UNIFICO LA CONSULTA
          
          Call APERBASDAT("CABAN")
          Screen.MousePointer = 11
          SQLX$ = "SELECT * FROM CAJABANC "
          SQLX$ = SQLX$ + "WHERE TipoMovim = 'CHRECI' "
          SQLX$ = SQLX$ + "AND Status = 1 "
          SQLX$ = SQLX$ + "AND OPCIMOVIM = 'CH' "
          SQLX$ = SQLX$ + " AND DescriMov = '" & IDEVA$ & "'"
          SQLX$ = SQLX$ + " AND ABS(ValAbsComp - " & IMCHEX$ & ") < 0.005 "
          Dim CHETEMP1 As ADODB.Recordset
          Set CHETEMP1 = New ADODB.Recordset
25        On Error Resume Next
          CHETEMP1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText

          If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
             On Error GoTo 0
             Call CapturaErrorOpen
             GoTo 25
          End If
          On Error GoTo 0

          With CHETEMP1
             On Error Resume Next
             '.FindFirst SQLX$
             'If .NoMatch = False Then
             .MoveFirst
             If Err < 1 Then
                On Error GoTo 0
'                .Edit
                !FEDEPOSENT = CVDAT(FF%)
                !Document = NRECIB$
                !DESTICHEQ = rasocli$((-1) * NC%)
                  CUECONTABINT% = !CUECONTAI  ' PARA QUE LA SALIDA SEA DE LA MISMA CUENTA QUE ENTRO
                !REFESALCHE = "Orden de Pago Nro." + Str$(NRO&)
                !IDENSALCHE = AJUSTI$(HOS$ + " - " + Left$(Time$, 5) + " hs - Term: " + NOMTER$, 48)
                !USUASALCHE = AJUSTI$("Usuario: " + TRIM$(USERNAME$), 48)
                !Status = 2                     ' Marca Salida de Cartera
                !IdenBaja = IDACTU$             ' Marca Modificacion de registro - FECHA - HORA - USUARIO - TERMINAL
                .Update
                ENCUEN% = 1
             End If
             On Error GoTo 0
          End With
          '
          ' ESTO SI NO ESTÁ INGRESADO
          If ENCUEN% < 1 Then
             QUE% = CVI(Mid$(Z$, 45, 2))
             For UX% = 1 To CAMP%
                 If MEPA%(UX%) = QUE% Then GoTo 38
             Next UX%
             ' PONER AQUI UNA OPCIÓN POR DEFAULT
             UX% = 3   ' ESTO ES EFECTIVO EN FERVI
             '
38           CUECONTABINT% = CUEMP%(UX%)
          End If
          '
          NC1% = (-1) * NC%
          IMPOP# = IMPOP# - IMCHE#
          CoComCor$ = NRECIB$
          CoComLar$ = NROX$
'          CodiEmp% = 0
'          CodiSuc% = SUC%
          NumeClie = 0
          NumeProv% = NC%
          FechEmi = CVDAT(FECHE%)
            FEX = FF%
          Referen$ = rasocli$(NC1%) + " - " + FECHATEX$(FEX)
          FeContab = CVDAT(FF%)
          '
          FeAcredi = CVDAT(FEVEN1%)
          CuentaCo% = CUECONTABINT%
          TiMovi$ = "CHENTR"
          OpMovi$ = "OP"
          VaMovi$ = ""
          StatInit = 0
          Descrip$ = IDEVA$
          NumeTal = COBRACT%
          NumeCom = NRO&
          ImpoDeb# = (-1) * Abs(IMCHE#)
          BasImpo# = 0               ' CERO PARA LA SALIDA DEL CHEQUE
          Observa$ = "O.Pago " + TRIM$(Str$(NRO&)) + " - " + rasocli$(NC1%)
          '
          Call GRACABAN
           '
       End If
    Next KKUU&
    CHETEMP1.Close
    Set CHETEMP1 = Nothing
    '
   ' Grabar Vencimientos de Facturas ******************************************
39 SIGNO% = (-1)
   IBruTot# = Abs(ROUND(CVD(Mid$(XX$, 59, 8))))
   IIvaTot# = 0
   '
   PRICONVEN% = -1
   SALBRUTO# = Abs(IBruTot#) * SIGNO%
   TXVENCI$ = Mid$(XX$, 139, 60)
   For KKI% = 1 To 6
      FEVEN%(KKI%) = CVI(Mid$(TXVENCI$, 10 * KKI% - 9, 2))
      VENCIPESOS#(KKI%) = CVD(Mid$(TXVENCI$, 10 * KKI% - 7, 8))
   Next KKI%
   '
   For KKI% = 1 To 6
      If FEVEN(KKI%) <> 0 Then
        If Abs(VENCIPESOS#(KKI%)) >= 0.005 Then
           NuorVen% = KKI%
           FeVenc = CVDAT(FEVEN(KKI%))
           ImVenc# = VENCIPESOS#(KKI%)
           Call GRACREPEN
         End If
      End If
   Next KKI%
   '
Return

GRAPLIP:
  With Aplipago
    On Error Resume Next
1301   .MoveFirst
    If Err = 0 Then
      .Delete
      GoTo 1301
    End If
    On Error GoTo 0
  End With
  '
  FIN& = ULTREG&("APLIPAG")
  For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XX$ = REGLEIDO$("APLIPAG", U&)
     NumeProv% = CVI(Mid$(XX$, 3, 2))
     CoComLar$ = Mid$(XX$, 5, 18)
        NROX$ = CoComLar$
     CoComCor$ = Left$(NROX$, 3) + Mid$(NROX$, 9, 1) + "-" + Mid$(NROX$, 14)
     ImpApli# = CVD(Mid$(XX$, 23, 8))
     RefApli$ = Mid$(XX$, 31, 24)
     DocOrLar$ = Mid$(XX$, 55, 18)
        NROX$ = DocOrLar$
     DocOrCor$ = Left$(NROX$, 3) + Mid$(NROX$, 9, 1) + "-" + Mid$(NROX$, 14)
     TotOrig# = CVD(Mid$(XX$, 73, 8))
     IvaOrig# = CVD(Mid$(XX$, 81, 8))
     Call GRAPLIPAG
   Next U&
Return

End Sub

