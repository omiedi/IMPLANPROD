VERSION 5.00
Begin VB.Form MENUACRE04 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFE0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestión de Compras - Cuentas Corrientes"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   5055
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   5055
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00800000&
      Height          =   4425
      Left            =   4095
      ScaleHeight     =   4365
      ScaleWidth      =   915
      TabIndex        =   12
      Top             =   0
      Width           =   975
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   3150
         Width           =   640
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -420
            TabIndex        =   16
            Top             =   0
            Width           =   12000
         End
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
         Left            =   140
         Picture         =   "MENUACRE04.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Ayuda Flexoft en Línea"
         Top             =   900
         Width           =   650
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
         Left            =   140
         Picture         =   "MENUACRE04.frx":030A
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
         TabIndex        =   17
         Top             =   1785
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -290
         Picture         =   "MENUACRE04.frx":0454
         Top             =   3440
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFE0&
      Caption         =   "Gestión de Compras"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
         BackColor       =   &H00FFFFE0&
         Caption         =   "Control de Gestión"
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
         TabIndex        =   8
         Top             =   2310
         Width           =   3480
         Begin VB.PictureBox Picture3 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   420
            ScaleHeight     =   585
            ScaleWidth      =   2580
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
               Height          =   600
               Left            =   0
               Picture         =   "MENUACRE04.frx":2376
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
               TabIndex        =   11
               Top             =   105
               Width           =   1905
            End
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFE0&
         Caption         =   "Operaciones Comerciales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1905
         Left            =   210
         TabIndex        =   3
         Top             =   315
         Width           =   3480
         Begin VB.PictureBox Picture8 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   735
            ScaleHeight     =   585
            ScaleWidth      =   2265
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
               Height          =   600
               Left            =   0
               Picture         =   "MENUACRE04.frx":2680
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
               TabIndex        =   5
               Top             =   105
               Width           =   1485
            End
         End
         Begin VB.PictureBox Picture14 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   420
            ScaleHeight     =   585
            ScaleWidth      =   2265
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
               Height          =   600
               Left            =   0
               Picture         =   "MENUACRE04.frx":27CA
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
               Top             =   60
               Width           =   1590
            End
         End
      End
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   5880
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu op_comer 
      Caption         =   "&Operaciones Comerciales"
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
         Caption         =   "Terminar"
      End
   End
   Begin VB.Menu cons_inf 
      Caption         =   "&Consultas e Informes"
      Begin VB.Menu cta_cte 
         Caption         =   "Resumen Cuenta Corriente"
         Begin VB.Menu scr1 
            Caption         =   "Individual por pantalla"
         End
         Begin VB.Menu print1 
            Caption         =   "Imprimir / Exportar"
         End
      End
      Begin VB.Menu cta_porpagar 
         Caption         =   "Cuentas por Pagar"
      End
      Begin VB.Menu plh5 
         Caption         =   "-"
      End
      Begin VB.Menu saldos 
         Caption         =   "Saldos de Proveedores"
         Begin VB.Menu scr2 
            Caption         =   "Por Pantalla"
         End
         Begin VB.Menu print2 
            Caption         =   "Imprimir / Exportar"
         End
      End
      Begin VB.Menu res_men 
         Caption         =   "Resúmenes Mensuales"
         Begin VB.Menu iva 
            Caption         =   "Subdiarios"
            Begin VB.Menu sub_ivacom 
               Caption         =   "Subdiario de IVA Compras"
            End
            Begin VB.Menu sub_compras 
               Caption         =   "Subdiario de Compras"
            End
         End
         Begin VB.Menu cons_tipo 
            Caption         =   "Detalle por Tipo de Comprobante"
            Begin VB.Menu ctipo_fac 
               Caption         =   "Facturación de Compras"
            End
            Begin VB.Menu ctipo_ndeb 
               Caption         =   "Notas de Débito"
            End
            Begin VB.Menu ctipo_ncre 
               Caption         =   "Notas de Crédito"
            End
            Begin VB.Menu plh6 
               Caption         =   "-"
            End
            Begin VB.Menu ctipo_opago 
               Caption         =   "Ordenes de Pago"
            End
         End
         Begin VB.Menu plh7 
            Caption         =   "-"
         End
         Begin VB.Menu res_porcta 
            Caption         =   "Resumen Mensual por Cuenta"
         End
      End
   End
   Begin VB.Menu help 
      Caption         =   "A&yuda"
   End
End
Attribute VB_Name = "MENUACRE04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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
   If DERACCE%("ORDEPAGO", "Ordenes de Pagos") >= 2 Then
      'ORPAGO04.Show 1
      Call RECLAPAG
      ORPAGONUE.Show 1
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
    FORFAC04.Show 1
    Command22.Enabled = True
    '
End Sub

Private Sub Command29_Click()
   Call HELPONLINE("ACRE_GES")
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   COGECOM04.Show 1
   Command3.Enabled = True
End Sub

Private Sub cta_porpagar_Click()
   Call COGECOM04.Command3_Click
End Sub

Private Sub ctipo_fac_Click()
   Call FRESHALL
   Call CARSELMESA
   Call SELECHO("SELMESA")
   MESASEL$ = VALACT1$("SELMESA")
   If TRIM$(MESASEL$) <> "" Then
     REMECO04.MESAN = MESASEL$
     Call REMECO04.Command3_Click
   End If
End Sub

Private Sub ctipo_ncre_Click()
   Call FRESHALL
   Call CARSELMESA
   Call SELECHO("SELMESA")
   MESASEL$ = VALACT1$("SELMESA")
   If TRIM$(MESASEL$) <> "" Then
     REMECO04.MESAN = MESASEL$
     Call REMECO04.Command5_Click
   End If
End Sub

Private Sub ctipo_ndeb_Click()
   Call FRESHALL
   Call CARSELMESA
   Call SELECHO("SELMESA")
   MESASEL$ = VALACT1$("SELMESA")
   If TRIM$(MESASEL$) <> "" Then
     REMECO04.MESAN = MESASEL$
     Call REMECO04.Command4_Click
   End If
End Sub

Private Sub ctipo_opago_Click()
   Call FRESHALL
   Call CARSELMESA
   Call SELECHO("SELMESA")
   MESASEL$ = VALACT1$("SELMESA")
   If TRIM$(MESASEL$) <> "" Then
     REMECO04.MESAN = MESASEL$
     Call REMECO04.Command2_Click
   End If
End Sub

Private Sub exit_Click()
   Call Command1_Click
End Sub

Private Sub fac_provee_Click()
   Call Command22_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     OPENFORMS = DoEvents
   Wend
End Sub

Private Sub Form_Load()
   '
   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then Call FINAL("")
   '
   EJER% = 1
   '
   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
      Caption = Caption + "  -  " + EPAC$
   End If
   Call GRATITFOR(Caption)
   '
   If Left$(APLINVO$, 7) = "FACTURA" Then
       FORFAC04.Show 1
       Call FINAL("")
     ElseIf Left$(APLINVO$, 6) = "ORPAGO" Then
       ORPAGONUE.Show 1
       Call FINAL("")
     ElseIf Left$(APLINVO$, 8) = "CGCOMPRA" Then
       COGECOM04.Show 1
       Call FINAL("")
     ElseIf Left$(APLINVO$, 8) = "SUBDESTA" Then
       REMECO04.Show 1
       Call FINAL("")
       
   End If

   If InStr(UCase$(EPAC$), "FERVI") > 0 Then
      HOII% = HOY(HOS$)
      If HOII% <= FECHANUM("28/11/06") Then
         If TRIM$(CONTROL$("MIGRAFER", "CHECARTE")) = "" Then
            Show
            Call RECUCHECAR
            Call FINAL("")
         End If
      End If
   End If
   '
10 CUMA$ = CONTROL$("AMAPRO", "ECUECON")
   If CUMA$ = "" Or CUEINT%(CUMA$) < 1 Then
     Call MENSERR(24, "Falta Cuenta Contable por Omisión\Correspondiente a 'Acreedores Comerciales'`.\  \Establezca la Cuenta Antes de Continuar.")
     Call CIERRARCH("*.*")
     Call CONECRUN("ARCHIG04/OPMASPRO", "Archivos Maestros")
     Call FINAL("")
   End If
   '
'   If Left$(APLINVO$, 6) = "ORPAGO" Then
'       ORPAGO04.Show 1
'       Call FINAL("")
'     ElseIf Left$(APLINVO$, 6) = "FACTURA" Then
'       FORFAC04.Show 1
'       Call FINAL("")
'     ElseIf Left$(APLINVO$, 6) = "CGCOMPRAS" Then
'       COGECOM04.Show 1
'       Call FINAL("")
'   End If
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
   Call Command29_Click
End Sub

Private Sub orpago_Click()
   Call Command14_Click
End Sub

Private Sub print1_Click()
   Call RECUEP04.Command3_Click
End Sub

Private Sub print2_Click()
   COGECOM04.Option2.Value = True
   Call COGECOM04.Command5_Click
End Sub

Private Sub res_porcta_Click()
   Call FRESHALL
   Call CARSELMESA
   Call SELECHO("SELMESA")
   MESASEL$ = VALACT1$("SELMESA")
   If TRIM$(MESASEL$) <> "" Then
     REMECO04.MESAN = MESASEL$
     Call REMECO04.Command6_Click
   End If
End Sub

Private Sub scr1_Click()
   Call COGECOM04.Command7_Click
End Sub

Private Sub scr2_Click()
   COGECOM04.Option1.Value = True
   Call COGECOM04.Command5_Click
End Sub

Private Sub sub_compras_Click()
   Call FRESHALL
   Call CARSELMESA
   Call SELECHO("SELMESA")
   MESASEL$ = VALACT1$("SELMESA")
   If TRIM$(MESASEL$) <> "" Then
     REMECO04.MESAN = MESASEL$
     Call REMECO04.Command16_Click
   End If
End Sub

Private Sub sub_ivacom_Click()
   Call FRESHALL
   Call CARSELMESA
   Call SELECHO("SELMESA")
   MESASEL$ = VALACT1$("SELMESA")
   If TRIM$(MESASEL$) <> "" Then
     REMECO04.MESAN = MESASEL$
     Call REMECO04.Command21_Click
   End If
End Sub

Sub RECUCHECAR()
    '
    Call COMUNI("FLEXOFT Debe Controlar y Quizá Recuperar\la Cartera de Cheques y Valores.\   \El Proceso Puede Demandar Varios Minutos.\  \No Utilizar las Aplicaciones\de Cuentas Corrientes Acreedoras hasta Finalizar.\  \Pulse 'Aceptar Mensaje' para Comenzar.")
    '
100 'RECUPERACION DE MOVIMIENTOS ANULADOS
   Call APERBASDAT("CABAN")
   With CajaBanc
      On Error Resume Next
      .MoveLast
      If Err Then
         On Error GoTo 0
         GoTo 200
      End If
      FIN& = .RecordCount
      JJ& = 0
      .MoveFirst
      Do While Not .EOF
         JJ& = JJ& + 1
         Call TRACE(20, JJ&, FIN&)
         If Mid$(!CodiCom_Lar, 6, 4) = "anul" Then
            If Abs(!ValAbsComp) > 0 Then
               .Edit
               !ValAbsComp = 0
               !Idebecomp = 0
               !IHABECOMP = 0
               !BASEIMPO = 0
               .Update
            End If
         End If
      .MoveNext
      Loop
   End With
   '
200 ' ESTO ES PARA RECUPERAR MOVIMIENTOS 'AS' EN FERVI
   '
   Call APERBASDAT("CABAN")
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + "WHERE NuProve > 0 "
   SQLX$ = SQLX$ + "AND CodiCom_Lar like 'AS-*' "
   Set Cabasel1 = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
   '
   With Cabasel1
     On Error Resume Next
130    .MoveFirst
       If Err = 0 Then
         .Delete
         GoTo 130
       End If
     On Error GoTo 0
   End With
   '
   SQLX$ = "SELECT * FROM SACREPEN "
   SQLX$ = SQLX$ + "WHERE NuProve > 0 "
   SQLX$ = SQLX$ + "AND CodiCom_Lar like 'AS-*' OR CodiCom_Lar like 'FF-*' "
   Set Cabasel1 = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
   '
   With Cabasel1
     On Error Resume Next
140    .MoveFirst
       If Err = 0 Then
         .Delete
         GoTo 140
       End If
     On Error GoTo 0
   End With
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
   Set CHETEMP1 = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
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
         .FindFirst "DescriMov = '" & IDEVA$ & "' "
         If .NoMatch = False Then
              Call REPLA(XX$, "**", 255, 2)
              Call GRAREG("BADACHE", XX$, U&)
              GoTo 19    ' LO ENCONTRO
            Else
              GoTo 29
         End If
         '
19       If FENTREGA% > 0 Then
            ' MARCAR LA SALIDA
            .Edit
            !Status = 2
            !FEDEPOSENT = CVDAT(FENTREGA%)
            !FECHENTRE = CVDAT(FENTREGA%)
            If TRIM$(" " & !Document) = "" Then !Document = "Mig.Datos"
            If TRIM$(" " & !DESTICHEQ) = "" Then !DESTICHEQ = DESTI$
            If TRIM$(" " & !REFESALCHE) = "" Then !REFESALCHE = OB1$
            If TRIM$(" " & !IDENSALCHE) = "" Then !IDENSALCHE = OB2$ ' AJUSTI$(HOS$ + " - " + Left$(Time$, 5) + " hs - Term: " + NOMTER$, 48)
            If TRIM$(" " & !UsuaSalChe) = "" Then !UsuaSalChe = OB3$ ' AJUSTI$("Usuario: " + TRIM$(USERNAME$), 48)
            .Update
         End If
      End With
      '
29 Next U&
   '
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + "WHERE TipoMovim = 'CHRECI'"
   SQLX$ = SQLX$ + "AND Status = 1"
   Set CHETEMP1 = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
   '
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
         IDEVA$ = !Descrimov
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
                   .Edit
                   !Status = 1
                   !FEDEPOSENT = 0
                   !FECHENTRE = 0
                   !Document = ""
                   !DESTICHEQ = DESTI$
                   !REFESALCHE = OB1$
                   !IDENSALCHE = OB2$ ' AJUSTI$(HOS$ + " - " + Left$(Time$, 5) + " hs - Term: " + NOMTER$, 48)
                   !UsuaSalChe = OB3$ ' AJUSTI$("Usuario: " + TRIM$(USERNAME$), 48)
                   .Update
                 Else
                   .Edit
                   !Status = 2
                   !FEDEPOSENT = CVDAT(FENTREGA%)
                   !FECHENTRE = CVDAT(FENTREGA%)
                   If TRIM$(" " & !Document) = "" Then !Document = "Mig.Datos"
                   If TRIM$(" " & !DESTICHEQ) = "" Then !DESTICHEQ = DESTI$
                   If TRIM$(" " & !REFESALCHE) = "" Then !REFESALCHE = OB1$
                   If TRIM$(" " & !IDENSALCHE) = "" Then !IDENSALCHE = OB2$ ' AJUSTI$(HOS$ + " - " + Left$(Time$, 5) + " hs - Term: " + NOMTER$, 48)
                   If TRIM$(" " & !UsuaSalChe) = "" Then !UsuaSalChe = OB3$ ' AJUSTI$("Usuario: " + TRIM$(USERNAME$), 48)
                   .Update
              End If
              GoTo 39
           End If
         Next U&
         '
         ' NO LO ENCONTRO EN BADACHE
         .Edit
         !Status = 2
         !FEDEPOSENT = !FechConta
         !FECHENTRE = !FechConta
         If TRIM$(" " & !Document) = "" Then !Document = "Mig.Datos"
         If TRIM$(" " & !DESTICHEQ) = "" Then !DESTICHEQ = "Nunca Entró"
         If TRIM$(" " & !REFESALCHE) = "" Then !REFESALCHE = "No Figurado en BADACHE"
         If TRIM$(" " & !IDENSALCHE) = "" Then !IDENSALCHE = AJUSTI$(HOS$ + " - " + Left$(Time$, 5) + " hs - Term: " + NOMTER$, 48)
         If TRIM$(" " & !UsuaSalChe) = "" Then !UsuaSalChe = AJUSTI$("Usuario: " + TRIM$(USERNAME$), 48)
         .Update
39    .MoveNext
      Loop
      '
   End With
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
   For I& = 1 To ULTREG&("MEPAGO")
       F0$ = REGLEIDO$("MEPAGO", I&)
       CAMP% = I&
       MEPA%(CAMP%) = CVI(Left$(F0$, 2))
       CUEX$ = Mid$(F0$, 45, 12)
       CUEMP%(CAMP%) = CUEINT%(CUEX$)
       If TRIM$(CODCUE$(CUEMP%(CAMP%))) <> TRIM$(CUEX$) Then
          Call MENSERR(24, "Cuenta Contable Incorrecta\para Medio de Pago '" + TRIM$(Mid$(F0$, 3, 12)) + "'\-\Resultados Imprevisibles.")
          Call FINAL("")
       End If
   Next I&
   '
   Call APERBASDAT("CABAN")
   Screen.MousePointer = 11
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + "WHERE TipoMovim = 'CHRECI'"
   SQLX$ = SQLX$ + "AND Status = 1 "
   SQLX$ = SQLX$ + "AND OPCIMOVIM = 'CH' "
   Set CHETEMP1 = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
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
   CodiEmp% = 0
   CodiSuc% = SUC%
   NumeClie% = 0
   NumeProv% = NC%
   FechEmi = CVDAT(FF%)
   '
   Referen$ = RASOCLI$(NC1%)
   FeContab = CVDAT(FF%)
   FeAcredi = CVDAT(FF%)
   NumeTal = 0
   NumeCom = NRO&
      If Observa$ = "" Then Observa$ = RASOCLI$(NC1%)
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
         Descrip$ = TRIM$(DEMOVX$) + " - " + RASOCLI$(NC1%)
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
     CuentaCo% = CUECOINT%(CONTROL$("", "CUEIVA-0"))
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
     CuentaCo% = CUECOINT%(CONTROL$("", "CUEIVA-0"))
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
         CuentaCo% = CUECOINT%(CONTROL$("", "CUEIVA-0"))
       ElseIf KKI% = 83 Then
         VaMovi$ = "Aduana"
         CuentaCo% = CUECOINT%(CONTROL$("", "CUEIVA-4"))
       ElseIf KKI% = 91 Then
         VaMovi$ = "Per-IVA"
         Descrip$ = "Percepcion I.V.A. " + NRECIB$
         CuentaCo% = CUECOINT%(CONTROL$("", "CUEIVA-5"))
       ElseIf KKI% = 99 Then
         TiMovi$ = "GAN"
         VaMovi$ = "Per-Gan"
         Descrip$ = "Percepcion Ganancias " + NRECIB$
         CuentaCo% = CUECOINT%(CONTROL$("", "CUEIVA-6"))
       ElseIf KKI% = 107 Then
         TiMovi$ = "I-BRU"
         VaMovi$ = "PERC"
         Descrip$ = "Percepcion I.Brutos " + NRECIB$
' ojo que esto está mal
         CuentaCo% = CUECOINT%(CONTROL$("", "CUEIVA-6"))
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
   Descrip$ = CoComLar$ + " - " + RASOCLI$(NC1%)
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
    CodiEmp% = 0
    CodiSuc% = SUC%
    NumeClie% = 0
    NumeProv% = NC%
    FechEmi = CVDAT(FF%)
    Referen$ = RASOCLI$((-1) * NC%)
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
    If Observa$ = "" Then Observa$ = RASOCLI$(NC1%)
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
          SQLX$ = "DescriMov = '" & IDEVA$ & "' "
          SQLX$ = SQLX$ + "AND ABS(ValAbsComp - " & IMCHEX$ & ") < 0.005 "
          SQLX$ = SQLX$ + "AND Status = 1"
          '
          With CHETEMP1
             .FindFirst SQLX$
             If .NoMatch = False Then
                On Error GoTo 0
                .Edit
                !FEDEPOSENT = CVDAT(FF%)
                !Document = NRECIB$
                !DESTICHEQ = RASOCLI$((-1) * NC%)
                  CUECONTABINT% = !CueContaI  ' PARA QUE LA SALIDA SEA DE LA MISMA CUENTA QUE ENTRO
                !REFESALCHE = "Orden de Pago Nro." + Str$(NRO&)
                !IDENSALCHE = AJUSTI$(HOS$ + " - " + Left$(Time$, 5) + " hs - Term: " + NOMTER$, 48)
                !UsuaSalChe = AJUSTI$("Usuario: " + TRIM$(USERNAME$), 48)
                !Status = 2                     ' Marca Salida de Cartera
                !IdenBaja = IDACTU$             ' Marca Modificacion de registro - FECHA - HORA - USUARIO - TERMINAL
                .Update
                ENCUEN% = 1
             End If
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
          CodiEmp% = 0
          CodiSuc% = SUC%
          NumeClie% = 0
          NumeProv% = NC%
          FechEmi = CVDAT(FECHE%)
            FEX = FF%
          Referen$ = RASOCLI$(NC1%) + " - " + FECHATEX$(FEX)
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
          Observa$ = "O.Pago " + TRIM$(Str$(NRO&)) + " - " + RASOCLI$(NC1%)
          '
          Call GRACABAN
           '
       End If
    Next KKUU&
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
  With ApliPago
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

