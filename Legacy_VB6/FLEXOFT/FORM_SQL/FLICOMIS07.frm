VERSION 5.00
Begin VB.Form FLICOMIS07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Liquidación de Comisiones por Venta y Cobranza"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10815
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   10815
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture6 
      BackColor       =   &H00404000&
      Height          =   6945
      Left            =   9870
      ScaleHeight     =   6885
      ScaleWidth      =   1635
      TabIndex        =   15
      Top             =   0
      Width           =   1695
      Begin VB.CommandButton Command4 
         Height          =   540
         Left            =   120
         Picture         =   "FLICOMIS07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Borrar Liquidaciones Previas"
         Top             =   1400
         Width           =   650
      End
      Begin VB.CommandButton Command5 
         Height          =   540
         Left            =   120
         Picture         =   "FLICOMIS07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Opciones ... Cuenta Corriente Vendedores"
         Top             =   2480
         Width           =   650
      End
      Begin VB.CommandButton Command3 
         Height          =   540
         Left            =   120
         Picture         =   "FLICOMIS07.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1940
         Width           =   650
      End
      Begin VB.CommandButton Command2 
         Height          =   540
         Left            =   120
         Picture         =   "FLICOMIS07.frx":0C7E
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   840
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   17
         Top             =   3150
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   18
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command1 
         Height          =   540
         Left            =   120
         Picture         =   "FLICOMIS07.frx":0F88
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   315
         Width           =   650
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -300
         Picture         =   "FLICOMIS07.frx":10D2
         Top             =   3450
         Width           =   1515
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Ultima Liquidación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   210
      TabIndex        =   11
      Top             =   210
      Width           =   4635
      Begin VB.TextBox MESAN 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1365
         TabIndex        =   12
         Text            =   " "
         Top             =   525
         Width           =   2535
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1365
         TabIndex        =   21
         Top             =   1120
         Width           =   3165
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Modalidad:"
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
         Index           =   13
         Left            =   -975
         TabIndex        =   14
         Top             =   1155
         Width           =   2220
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Período:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   225
         Index           =   10
         Left            =   -945
         TabIndex        =   13
         Top             =   630
         Width           =   2220
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Modalidad de Cálculo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   5040
      TabIndex        =   7
      Top             =   210
      Width           =   4635
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Por Facturas CANCELADAS (Pago Total)"
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
         Left            =   420
         TabIndex        =   22
         Top             =   900
         Value           =   -1  'True
         Width           =   4110
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Por Cobranza Neta de I.V.A"
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
         Left            =   420
         TabIndex        =   10
         Top             =   1150
         Width           =   3375
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Por Facturación Neta del Periodo"
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
         Left            =   420
         TabIndex        =   9
         Top             =   650
         Width           =   3375
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Por Pedidos Netos y Cobranza"
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
         Left            =   420
         TabIndex        =   8
         Top             =   400
         Width           =   4005
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Totales por Vendedor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      Left            =   5040
      TabIndex        =   5
      Top             =   1785
      Width           =   4635
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1410
         Left            =   210
         TabIndex        =   6
         Top             =   380
         Width           =   4215
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Avance de Proceso"
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
      Left            =   210
      TabIndex        =   0
      Top             =   1970
      Width           =   4635
      Begin VB.CheckBox Check4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Cálculo de Comisiones"
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
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1365
         Width           =   2745
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Valorización por Producto"
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
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1050
         Width           =   2745
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Detalle de Facturas por Recibo"
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
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   735
         Width           =   3690
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Selección de Cobranza del Período"
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
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   420
         Width           =   3585
      End
   End
End
Attribute VB_Name = "FLICOMIS07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MODOCAL%
'
Dim SUPED#(256), COMPED#(256), SUCOB#(256), COMCOB#(256)
Dim PRIMOSTO&(8192), ULTMOSTO&(8192)
Dim IFAC#(8192), IDET#(8192), SIFAC%(8192)
Dim NRECI&(8192), REGFA$(8192)
Dim NCV%(8192), TII%(8192), NROV&(8192), FFII%(8192) ' dimensionamiento
'
Private Sub Check1_Click()
    Check1.Value = 0
End Sub

Private Sub Check2_Click()
    Check2.Value = 0
End Sub

Private Sub Check3_Click()
    Check3.Value = 0
End Sub

Private Sub Check4_Click()
    Check4.Value = 0
End Sub

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    Command3.Enabled = False
    YALIQUI% = 0
    MODOCAL% = 0
    If Option1.Value = True Then MODOCAL% = 1
    If Option3.Value = True Then MODOCAL% = 2
    If Option4.Value = True Then MODOCAL% = 3
    If Option5.Value = True Then MODOCAL% = 4
    If MODOCAL% < 1 Then
        Call MENSERR(24, "Falta Seleccionar Modalidad de Cálculo")
        GoTo 99
    End If
    If MODOCAL% = 2 Or MODOCAL% = 1 Then
       Call MENSERR(24, "Modalidad no Disponible en la Presente Versión.")
       GoTo 99
    End If
    '
    If MODOCAL% = 1 Then
         Call CALCSILCAR
       Else
         If Option5.Value = True Then
            Call COMUNI("Atención!!!\La Comision Aplicada se toma\ de la base de Datos del Vendedor")
         End If
         Call CALCOMIS
    End If
    '
99  Command2.Enabled = True
    If YALIQUI% = 1 Then
        Command3.Enabled = True
    End If
    '
End Sub

Private Sub Command3_Click()
    Call PRILICOMIS
End Sub

Sub PAGOVEN()
   '
   '\\\OT-007-0076-RG-28/02/07///
5  VDEF$ = REGBLAN$("CUECOMIS")
   HOII% = HOY(HOS$)
   Call REPLA(VDEF$, MKI$(HOII%), 2, 2)
   Call REPLA(VDEF$, Chr$(2), 4, 1)
   '
10 OBX$ = OBJPLA$("PAGOCOMIS", VDEF$)
   If OBX$ <> "" Then
     VENDE% = Asc(Left$(OBX$, 1))
     If VENDE% < 1 Or ECOARCH("VENDEDOR", Chr$(VENDE%)) = "" Then
       VDEF$ = OBX$
       Call MENSERR(24, "Falta Vendedor")
       GoTo 10
     End If
     '
     REFE$ = TRIM$(Mid$(OBX$, 9, 32))
     If REFE$ = "" Then
       VDEF$ = OBX$
       Call MENSERR(24, "Falta Referencia")
       GoTo 10
     End If
     '
     IMPO# = CVD(Mid$(OBX$, 49, 8))
     If Abs(IMPO#) < 0.005 Then
       VDEF$ = OBX$
       Call MENSERR(24, "Falta Importe")
       GoTo 10
     End If
     '
     If IMPO# < 0 Then
       Call REPLA(VDEF$, MKD$(Abs(IMPO#)), 41, 8)
       Call REPLA(VDEF$, MKD$(0), 49, 8)
     End If
     '
     Call REGAPP("CUECOMIS", OBX$)
     Call CIERRARCH("CUECOMIS")
     '
     Call COMUNI("Pago Registrado !!!")
     GoTo 5
     '
   End If
   '\\\OT-007-0076-RG-FIN///
   '
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   Call BORRACOMIS
   Command4.Enabled = True
End Sub

'
Private Sub Command5_Click()
    '
    '\\\OT-007-0076-RG-28/02/07///
    Command5.Enabled = False
    '\\\OT-007-0393-FL-19/10/07///
10  AMODO% = ALTERNA%("Cuenta Corriente Vendedores\Informar Pagos\Saldos Pendientes por Vendedor\Listado de Control Mensual\Envío de Comisiones por Mail\Modificar Comisiones de Facturas")
    If AMODO% = 1 Then
          OPC% = COMALTER%("Opciones de Trabajo:\\Consulta por Vendedor\Listado por Fechas")
          If OPC% = 1 Then
                Call CONECRUN("*CUECOMIS", "Consulta por Vendedor")
            ElseIf OPC% = 2 Then
                Call LICUECVEN
            ElseIf OPC% = 0 Then GoTo 10
          End If
    '\\\OT-007-0393-FL-FIN///
       ElseIf AMODO% = 2 Then
          Call PAGOVEN
          Call CONECRUN("*CUECOMIS", "Cuenta Corriente Vendedores")
          GoTo 10
       ElseIf AMODO% = 3 Then
          Call SALCOMI
       '28/03/07 (OT-07-0134)
       ElseIf AMODO% = 4 Then
          EPAC$ = TRIM$(UCase$(EMPREAC$))
          If InStr(EPAC$, "SILCAR") < 1 Then
            If InStr(EPAC$, "TELSO") < 1 Then
              Call MENSERR(24, "No Disponible en Esta Versión.")
              GoTo 99
            End If
          End If
          Call CONTROLCOMIS
       '28/03/07 (FIN)
       ElseIf AMODO% = 5 Then
         Call CONECRUN("SNDCOM07", "Envío de Comisiones por Mail")
       ElseIf AMODO% = 6 Then
         Call ActuComision
    End If
99  Command5.Enabled = True
    '\\\OT-007-0076-RG-FIN///
    '
End Sub

Sub ActuComision()
    '
    Call ABRECONEXION
10  Call DESHASFECHA(DES%, HAS%, PERIO$)

    If PERIO$ = "" Then Exit Sub
    DESDAT = FETEXCOR1$(DES%)
    HASDAT = FETEXCOR1$(HAS%)
    '
    CONDI$ = "FechEmisi >= '" & DESDAT & "' "
    CONDI$ = CONDI$ + " AND FechEmisi <= '" & HASDAT & "' "
    CONDI$ = CONDI$ + " AND TipoMovim = 'FVEN'"
    CONDI$ = CONDI$ + " AND ValAbsComp > 0"
    '
    CANREG& = CantRegistros&("CAJABANC", CONDI$)
    '
    If CANREG& < 1 Then
       Call MENSERR(24, "No se han Encontrado Comprobantes en el Rango Seleccionado")
       GoTo 10
    End If
    '
    ' Call COPYSTRU("SELCUERE", "SELEFACN") ' LISTADO DE SELECCION
    '
    Call CARGALISEL("SELFAC", "CAJABANC", "CODICOM_LAR/A20;FechEmisi/D8;NuClien/I5+DEUDOR1/A20", CONDI$, "DISTINCT")
    Call FRESHECHO("SELFAC")
25  Call SELECHO("SELFAC/Facturas del Periodo")
    '
    NFACT$ = VALACT1$("SELFAC")
    '
    If NFACT$ = "" Then GoTo 10
    '
    CONDI1$ = " CODICOM_LAR = '" & NFACT$ & "'"
    '
    NC% = XVALO(VALOBADA("CAJABANC", "NUCLIEN", CONDI1$, "NOMESS"))
    RazonSoc$ = rasocli$(NC%)
    fech = VALOBADA("CAJABANC", "FechEmisi", CONDI1$, "NOMESS")
    '
    VEND% = VENDCLI%(NC%)
    FIN& = ULTREG&("VENDEDOR")
    For i& = 1 To FIN&
        TX$ = REGLEIDO$("VENDEDOR", i&)
        If Asc(Mid$(TX$, 1, 1)) = VEND% Then
           COMIS = CVS(Mid$(TX$, 44, 4))
           Exit For
        End If
    Next i&
    COMIFAC = XVALO(VALOBADA("CAJABANC", "LETRACOMIS", CONDI1$, "NOMESS"))
    If COMIFAC > 0 Then COMIS = COMIFAC
    '
    Call REPLA(VDEF$, MKI$(CVINT(fech)), 1, 2) ' INFORMACION PARA EL OBJETO
    Call REPLA(VDEF$, NFACT$, 3, 18)
    Call REPLA(VDEF$, MKI$(NC%), 21, 2)
    Call REPLA(VDEF$, RazonSoc$, 23, 30)
    Call REPLA(VDEF$, MKS(COMIS), 53, 4)
    '
30  OBX$ = OBJPLA$("ACTUCOMI", VDEF$) ' ABRE EL OBJETO
    '
    If OBX$ = "" Then GoTo 25
    '
    CMSION# = CVS(Mid$(OBX$, 53, 4))
    If CMSION# < 0.005 Or CMISION# > 99.99 Then
      Call COMUNI("Ingreso No Válido")
      VDEF$ = OBX$
      GoTo 30
    End If
    CMISION1# = ROUND#(CMSION#, 2)
    '
    OPC% = COMALTER%("Confirma Modificacion:\\No\Si")
    '
    If OPC% = 2 Then ' GRABACION
        FECHX = FETEXCOR1(CVINT(fech))
        CONDI2$ = " CODICOM_LAR = '" & NFACT$ & "' AND FechEmisi = '" & FECHX & "' AND NUCLIEN = " & NC% & ""
        Call ACTUREGISTRO("CAJABANC", "LETRACOMIS", CONDI2$, CMISION1#, REGAF&, "NOMESS")
        GoTo 25 ' VUELVE AL SELECHO
    End If
    GoTo 25

End Sub
Private Sub Form_Load()
    '
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    Command3.Enabled = False
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Private Sub List1_DblClick()
   Call PRILICOMIS
End Sub

Private Sub MESAN_GOTFOCUS()
    Command2.SetFocus
End Sub

Private Sub Option3_Click()
    List1.Clear
    MESAN = "": Label1 = ""
    Command3.Enabled = False
End Sub

Private Sub Option4_Click()
    List1.Clear
    MESAN = "": Label1 = ""
    Command3.Enabled = False
End Sub

Private Sub Option5_Click()
    List1.Clear
    MESAN = "": Label1 = ""
    Command3.Enabled = False
End Sub

Sub CALCOMIS()
    'CORREGIDO
    Dim PORCO(256), TOCOMI#(256), TOCOB#(256)
    Label1 = ""
    MESAN = ""
    '
    MODOCAL% = 0
    If Option3.Value = True Then MODOCAL% = 2
    If Option4.Value = True Then MODOCAL% = 3   ' POR FACTURACIÓN
    If Option5.Value = True Then MODOCAL% = 4   ' POR COBRANZA
    '
    For U& = 1 To ULTREG&("VENDEDOR")
      x$ = REGLEIDO$("VENDEDOR", U&)
      VENDE% = Asc(Left$(x$, 1))
      If VENDE% > 0 Then
        If VENDE% <= 256 Then
          If PORCO(VENDE%) < 0.005 Then
            If MODOCAL% = 3 Then
                PORCO(VENDE%) = CVS(Mid$(x$, 44, 4))
              ElseIf MODOCAL% = 4 Then
                PORCO(VENDE%) = CVS(Mid$(x$, 48, 4))
            End If
          End If
        End If
      End If
    Next U&
    Call CIERRARCH("VENDEDOR")
    '
    Call DESHASFECHA(DES%, HAS%, PERIO$)
    If PERIO$ = "" Then Exit Sub
    MESAN = PERIO$
    DESDAT = CDbl(CVDAT(DES%))
    HASDAT = CDbl(CVDAT(HAS%))
    
    '
    CUECOMI% = 0
    If EXISTE%(LUDAT$ + "CUECOMIS.RFS") Then
      CUECOMI% = 1
      Screen.MousePointer = 11
      Call BLOQARCH("CUECOMIS")
      FIN& = ULTREG&("CUECOMIS")
      JJ& = 0
      For U& = 1 To FIN&
        XX$ = REGLEIDO$("CUECOMIS", U&)
        FECO% = CVI(Mid$(XX$, 2, 2))
        TIMX% = Asc(Mid$(XX$, 4, 1))
        If FECO% < DES% Or FECO% > HAS% Or TIMX% = 2 Then
          JJ& = JJ& + 1
          Call GRAREG("CUECOMIS", XX$, JJ&)
        End If
      Next U&
      Call SETULTREG("CUECOMIS", JJ&)
      Call CIERRARCH("CUECOMIS")
      Screen.MousePointer = 1
    End If
    '
    Screen.MousePointer = 11
    If MODOCAL% = 3 Or MODOCAL% = 4 Then      ' POR FACTURACION / COBRANZA
        REBLA$ = REGBLAN$("LICOMIN")
        Call BLANARCH("AVAFACB/NOERR")
        Call APERBASDAT("CABAN")
        SQLX$ = "SELECT * FROM CAJABANC "
        SQLX$ = SQLX$ + "WHERE FechEmisi >= '" & FETEXCOR1$(DES%) & "' "
        SQLX$ = SQLX$ + "AND FechEmisi <= '" & FETEXCOR1$(HAS%) & "' "
        If MODOCAL% = 3 Then
             SQLX$ = SQLX$ + "AND TipoMovim = 'VENTA' "
           Else
             SQLX$ = SQLX$ + "AND TipoMovim = 'RCOB' "
        End If
        SQLX$ = SQLX$ + "Order by FechEmisi ASC, NumeComp ASC "
        '
        'Set AVANCETMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
       Dim AVANCETMP As ADODB.Recordset
       Set AVANCETMP = New ADODB.Recordset
       AVANCETMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic

        With AVANCETMP
          On Error Resume Next
          .MoveFirst
          If Err Then
            On Error GoTo 0
            Screen.MousePointer = 1
            If MODOCAL% = 3 Then
                 Call MENSERR(24, "Sin Facturación en Período Seleccionado")
               ElseIf MODOCAL% = 4 Then
                 Call MENSERR(24, "Sin Cobranza en Período Seleccionado")
            End If
            Exit Sub
          End If
          '
          JJ& = 0
          y$ = REBLA$
          '
          Do While Not .EOF
            If MODOCAL% = 4 Or (!OpciMovim = "FC" Or !OpciMovim = "NC") Then
               NC% = !NUCLIEN
               CLIE% = ""
               If NC% > 0 Then
                  NUMVEN% = VENDCLI%(NC%)
                  NRO& = !NumeComp
                  NETO# = !IHABECOMP - !IDEBECOMP
                  NUSUCUVE% = !CodiSucu
                  NNX$ = !CODICOM_LAR
                  FEDOC% = CVINT(!FECHEMISI)
                  '
                  DIVI = 1
                  If MODOCAL% = 4 Then    ' NETEAR DE IVA SI POR COBRANZA
                    ' NO CONSIDERAR EL IVA
                    PVCL% = PIVACLI%(NC%)
                    DIVI = 1 + ALIVA / 100
                    If PVCL% = 2 Then DIVI = 1 + 1.5 * ALIVA / 100
                    If PVCL% = 5 Or PVCL% = 7 Then DIVI = 1
                  End If
                  '
                  Z$ = REBLA$
                  Call REPLA(Z$, MKI$(NC%), 1, 2)
                  Call REPLA(Z$, MKS$(NRO&), 3, 4)
                  Call REPLA(Z$, MKI$(FEDOC%), 7, 2)
                  Call REPLA(Z$, MKD$(NETO#), 9, 8)
                  Call REPLA(Z$, MKI$(NC%), 17, 2)
                  Call REPLA(Z$, MKI$(NUMVEN%), 19, 2)
                  Call REPLA(Z$, MKS$(0), 21, 4)
                  '
                  VENDE% = NUMVEN%
                  If VENDE% >= 0 Then
                     If VENDE% <= 256 Then
                        Call REPLA(Z$, MKS$(PORCO(VENDE%)), 45, 4)
                        ICOMI# = ROUND(NETO# * PORCO(VENDE%) / (100 * DIVI))
                        '
                        COMIS = XVALO(!LETRACOMIS)
                        If COMIS > 0 Then ' SI TIENE COMISION POR FACTURA,TOMA ESTA
                           Call REPLA(Z$, MKS$(COMIS), 45, 4)
                           ICOMI# = ROUND(NETO# * COMIS / (100 * DIVI))
                        End If
                         '
                        Call REPLA(Z$, MKD$(ICOMI#), 49, 8)
                        TOCOMI#(VENDE%) = TOCOMI#(VENDE%) + ICOMI#
                        TOCOB#(VENDE%) = TOCOB#(VENDE%) + NETO#
                     End If
                  End If
                  '
                  JJ& = JJ& + 1
                  Call GRAREG("LICOMIN", Z$, JJ&)
                  '
                  If CUECOMI% = 1 Then
                    YCM$ = REGBLAN$("CUECOMIS")
                    Call REPLA(YCM$, Chr$(NUMVEN%), 1, 1)
                    Call REPLA(YCM$, MKI$(FEDOC%), 2, 2)
                    Call REPLA(YCM$, Chr$(1), 4, 1)
                    Call REPLA(YCM$, MKS$(NRO&), 5, 4)
                    Call REPLA(YCM$, rasocli$(NC%), 9, 32)
                    Call REPLA(YCM$, MKD$(ICOMI#), 41, 8)
                    Call REGAPP("CUECOMIS", YCM$)
                  End If
                  '
               End If
            End If
49        .MoveNext
          Loop
        End With
        '
        Call SETULTREG("LICOMIN", JJ&)
        Call CIERRARCH("*.*")
        Screen.MousePointer = 1
        Call COMUNI("Calculo Completo")
        '
    End If
    '
    Call CIERRARCH("LICOMIN")
    List1.Clear
    '
    For VENDE% = 0 To 255
      If TOCOB#(VENDE%) > 0 Then
        If VENDE% = 0 Then
             NOVEN$ = "DIRECCION"
           Else
             NOVEN$ = ECOARCH$("VENDEDOR", Chr$(VENDE%))
        End If
        TXT$ = Space$(48)
        Call REPLA(TXT$, CSTRING$(MKI$(VENDE%), 1, 3, 0, 2), 1, 4)
        Call REPLA(TXT$, NOVEN$, 6, 17)
        Call REPLA(TXT$, CSTRING$(MKD$(TOCOMI#(VENDE%)), 4, 10, 2, 2), 23, 10)
        List1.AddItem TXT$
      End If
    Next VENDE%
    Check4.Value = 1
    '
    If MODOCAL% = 2 Then
        Label1 = "Por Pedidos"
      ElseIf MODOCAL% = 3 Then
        Label1 = "Por Facturación Neta"
      ElseIf MODOCAL% = 4 Then
        Label1 = "Por Cobranza Bruta"
    End If
    '
    If List1.ListCount > 0 Then
      List1.ListIndex = 0
    End If
    List1.Refresh
    Command3.Enabled = True
    '
End Sub
'
Sub CONTROLCOMIS()
    '
    '28/03/07 (OT-07-0134)
    '
    MESAN = ""
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) <> "" Then
      PERIO$ = TRIM$(MESANO$(MESASEL$, DES%, HAS%))
    End If
    '
    If PERIO$ = "" Then Exit Sub
    MESAN = PERIO$
    DESDAT = CDbl(CVDAT(DES%))
    HASDAT = CDbl(CVDAT(HAS%))
    '
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE FechEmisi >= " & DESDAT & " "
    SQLX$ = SQLX$ + "AND FechEmisi <= " & HASDAT & " "
    SQLX$ = SQLX$ + "AND TipoMovim = 'VENTA' "
    SQLX$ = SQLX$ + " AND (OpciMovim = 'FC' Or OpciMovim = 'NC') "
    SQLX$ = SQLX$ + "Order by FechEmisi ASC, CodiCom_Lar ASC "
    '
    'Set AVANCETMP = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
    Dim AVANCETMP As ADODB.Recordset
    Set AVANCETMP = New ADODB.Recordset
    AVANCETMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic
    '
    With AVANCETMP
        '
        On Error Resume Next
        .MoveFirst
        If Err Then
          On Error GoTo 0
          Screen.MousePointer = 1
          Call MENSERR(24, "Ningun Registro Coincide con el Filtro SQL")
          Exit Sub
        End If
        '
        y$ = REGBLAN$("LISCONME")
        Do While Not .EOF
           NC% = !NUCLIEN
           If NC% > 0 Then
              FEDOC% = CVINT(!FECHEMISI)
              NRO$ = !CODICOM_LAR
              IMPORTE# = !IHABECOMP - !IDEBECOMP
              PER$ = ""
              On Error Resume Next
              PER$ = !PerLiComis
              NVEND$ = Chr(!VendeComis)
              VEND$ = ECOARCH$("VENDEDOR", NVEND$)
              On Error GoTo 0
              If PER$ = "" Then
                 PER$ = "Pend.Liquid."
                 VEND$ = ""
              End If
              IMPCOMIS# = 0
              On Error Resume Next
              IMPCOMIS# = !ImpoComis
              On Error GoTo 0
              '
              Call REPLA(y$, MKI$(FEDOC%), 1, 2)
              Call REPLA(y$, NRO$, 3, 24)
              Call REPLA(y$, MKI$(NC%), 27, 2)
              Call REPLA(y$, MKD$(IMPORTE#), 29, 8)
              Call REPLA(y$, PER$, 37, 20)
              Call REPLA(y$, MKD$(IMPCOMIS#), 57, 8)
              Call REPLA(y$, TRIM$(VEND$), 65, 30)
              i& = i& + 1
              Call GRAREG("LISCONME", y$, i&)
            End If
          '
          .MoveNext
          Loop
    End With
    '
    Call HEADERS("LISCONME", "")
    Call HEADERS("LISCONME", "Período: " + PERIO$)
    '
    Call SETULTREG("LISCONME", i&)
    Call CIERRARCH("LISCONME")
    '
    Call FINAL("*LISCONME")
    '
    '28/03/07 (FIN)
    '
End Sub
'
Sub CALCSILCAR()
    '
    OPC% = COMALTER%("¿Desea Incluir Notas de Débito?\\ No \ Si ") ' Incluye notas de debitos enel calculo de comisiones
    If OPC% = 2 Then
      Call SELECHO("SELFECHA")
      FCORTE0$ = TRIM$(VALACT1$("SELFECHA"))
      FCORTE1 = 0: If FCORTE0$ <> "" Then FCORTE1 = FECHANUM(FCORTE0$)
    End If
    '
    Dim PORCO(256), TOCOMI#(256), TOCOB#(256)
    Label1 = ""
    MESAN = ""
    '
    For U& = 1 To ULTREG&("VENDEDOR")
      x$ = REGLEIDO$("VENDEDOR", U&)
      VENDE% = Asc(Left$(x$, 1))
      If VENDE% > 0 Then
        If VENDE% <= 256 Then
          If PORCO(VENDE%) < 0.005 Then
            PORCO(VENDE%) = CVS(Mid$(x$, 44, 4))
            A = PORCO(VENDE%)
          
          End If
        End If
      End If
    Next U&
    Call CIERRARCH("VENDEDOR")
    '
    MESAN = ""
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) <> "" Then
      PERIO$ = TRIM$(MESANO$(MESASEL$, DES%, HAS%))
    End If
    '
    If PERIO$ = "" Then Exit Sub
    MESAN = PERIO$
    DESDAT = CDbl(CVDAT(DES%))
    HASDAT = CDbl(CVDAT(HAS%))
    '
    '\\\OT-007-0076-RG-28/02/07///
    PERIOX$ = TRIM$(PERIO$)

    If EXISTE%(LUDAT$ + "CUECOMIS.RFS") Then
       CUECOMI% = 1
       Screen.MousePointer = 11
       Call BLOQARCH("CUECOMIS")
       FIN& = ULTREG&("CUECOMIS")
       JJ& = 0
       For U& = 1 To FIN&
           XX$ = REGLEIDO$("CUECOMIS", U&)
           TIMX% = Asc(Mid$(XX$, 4, 1))
           PER$ = TRIM$(Mid$(XX$, 97, 16))
           ' REGENERA LA CUENTA CORRIENTE DE VENDEDORES (CUECOMIS)
           ' EXCLUYENDO LAS LIQUIDACIONES REALIZADAS EN EL MISMO
           ' PERIODO ELEGIDO (PERIOX$)
           If PER$ <> PERIOX$ Or TIMX% = 2 Then
              JJ& = JJ& + 1
              Call GRAREG("CUECOMIS", XX$, JJ&)
           End If
       Next U&
       Call SETULTREG("CUECOMIS", JJ&)
       Call CIERRARCH("CUECOMIS")
       Screen.MousePointer = 1
    End If
    '\\\OT-007-0076-RG-FIN///
    '
    Screen.MousePointer = 11
    REBLA$ = REGBLAN$("LICOMACC")
    Call BLANARCH("LICOMACC/NOERR")
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + " WHERE FechEmisi <= '" & FETEXCOR1$(HAS%) & "'"
    SQLX$ = SQLX$ + " AND TipoMovim = 'VENTA' "
    If OPC% = 1 Then SQLX$ = SQLX$ + " AND (OpciMovim = 'FC' Or OpciMovim = 'NC') " 'Si toma en cuenta las nota de debito no filtra por Fc ni NC
    'SQLX$ = SQLX$ + " AND (OpciMovim = 'FC' Or OpciMovim = 'NC') "
    SQLX$ = SQLX$ + "AND (PerLiComis Is Null OR PerLiComis = '' OR PerLiComis = '" & PERIO$ & "') "
    SQLX$ = SQLX$ + "Order by FechEmisi ASC, NumeComp ASC "
    '
    Dim AVANCETMP As ADODB.Recordset
    Set AVANCETMP = New ADODB.Recordset
    AVANCETMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic
    'Set AVANCETMP = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
    With AVANCETMP
       On Error Resume Next
       .MoveFirst
       If Err Then
          On Error GoTo 0
          Screen.MousePointer = 1
          Call MENSERR(24, "Sin Facturación en Período Seleccionado")
          Exit Sub
       Else
          FIS& = AVANCETMP.RecordCount
       End If
       '
' VERIFICAR QUE NO ESTÉ YA LIQUIDADO O QUE COINCIDA EL PERÍODO
       JJ& = 0
       '
       Do While Not .EOF
          NC% = XVALO(!NUCLIEN)
          If NC% > 0 Then
              NNX$ = SAFETEXT$(!CODICOM_LAR)
              '
              Call TRACE(20, JJ&, FIS&)
              SIGNO% = 1  'Si el comprobante es nota de debito asigna -1 a la variable SIGNO%
              If Mid$(NNX$, 1, 2) = "ND" Then
                If CVINT(!FECHEMISI) < FCORTE1 Then GoTo 49
                SIGNO% = (-1) 'Si SIGNO% = -1 el importe de la nota de debito es tomado como negativo
                If Abs(SALDEUPEN#(NNX$)) < 0.005 Then GoTo 49 ' para no tener en cuenta las notas de debitos canceladas
                GoTo 20
              End If
              ' verificar que no figure como pendiente
              If Abs(SALDEUPEN#(NNX$)) > 0.005 Then GoTo 49
              '
20            NUMVEN% = VENDCLI%(NC%)
              NRO& = XVALO(!NumeComp)
              'NETO# = !IHabeComp - !IDebeComp
              NETO# = (XVALO(!IHABECOMP) - XVALO(!IDEBECOMP)) * SIGNO% 'En caso de haber notas de debito, resta el importe
              NUSUCUVE% = XVALO(!CodiSucu)
              FEDOC% = CVINT(!FECHEMISI)
              '
              Z$ = REBLA$
              Call REPLA(Z$, MKI$(NC%), 1, 2)
              Call REPLA(Z$, NNX$, 3, 18)
              Call REPLA(Z$, MKI$(FEDOC%), 21, 2)
              Call REPLA(Z$, MKD$(NETO#), 23, 8)
              Call REPLA(Z$, MKI$(NC%), 31, 2)
              Call REPLA(Z$, MKI$(NUMVEN%), 33, 2)
              Call REPLA(Z$, MKS$(0), 35, 4)
              '
              VENDE% = NUMVEN%
              ICOMI# = 0
              If VENDE% >= 0 Then
                 If VENDE% <= 256 Then
                    '
                    Call REPLA(Z$, MKS$(PORCO(VENDE%)), 35, 4)
                    ICOMI# = ROUND(NETO# * PORCO(VENDE%) / 100)
                    '
                    COMIS = XVALO(!LETRACOMIS)
                    If COMIS > 0 Then ' SI TIENE COMISION POR FACTURA,TOMA ESTA
                       Call REPLA(Z$, MKS$(COMIS), 35, 4)
                       ICOMI# = ROUND(NETO# * COMIS / 100)
                    End If
                    '
                    Call REPLA(Z$, MKD$(ICOMI#), 39, 8)
                    TOCOMI#(VENDE%) = TOCOMI#(VENDE%) + ICOMI#
                    TOCOB#(VENDE%) = TOCOB#(VENDE%) + NETO#
                 End If
              End If
              '
              JJ& = JJ& + 1
              Call GRAREG("LICOMACC", Z$, JJ&)
              '\\\OT-007-0076-RG-28/02/07///
              'CALCULA LAS COMISIONES DEL MES Y LAS INCLUYE DENTRO
              'DE LA CUENTA CORRIENTE DEL VENDEDOR
              If CUECOMI% = 1 Then
                 YCM$ = REGBLAN$("CUECOMIS")
                 Call REPLA(YCM$, Chr$(VENDE%), 1, 1)
                 Call REPLA(YCM$, MKI$(HAS%), 2, 2)
                 Call REPLA(YCM$, Chr$(1), 4, 1)
                 Call REPLA(YCM$, MKS$(NRO&), 5, 4)
                 Call REPLA(YCM$, rasocli$(NC%), 9, 32)
                 Call REPLA(YCM$, MKD$(ICOMI#), 41, 8)
                 REFE$ = SAFETEXT$(!CODICOM_LAR) + " - " + FECHATEX$(FEDOC%)
                 Call REPLA(YCM$, REFE$, 65, 32)
                 Call REPLA(YCM$, PERIO$, 97, 16)
                 Call REGAPP("CUECOMIS", YCM$)
              End If
              '\\\OT-007-0076-RG-FIN///
              '.Edit
              !PerLiComis = PERIO$
              !ImpoComis = ICOMI#
              !VendeComis = NUMVEN%
              .Update
              '
          End If
49     .MoveNext
       Loop
    End With
    '
    Call SETULTREG("LICOMACC", JJ&)
    Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    Call COMUNI("Calculo Completo")
    '
    Call CIERRARCH("LICOMACC")
    List1.Clear
    '
    For VENDE% = 0 To 255
      If Abs(TOCOB#(VENDE%)) >= 0.005 Then
        If VENDE% = 0 Then
             NOVEN$ = "DIRECCION"
           Else
             NOVEN$ = ECOARCH$("VENDEDOR", Chr$(VENDE%))
        End If
        TXT$ = Space$(48)
        Call REPLA(TXT$, CSTRING$(MKI$(VENDE%), 1, 3, 0, 2), 1, 4)
        Call REPLA(TXT$, NOVEN$, 6, 17)
        Call REPLA(TXT$, CSTRING$(MKD$(TOCOMI#(VENDE%)), 4, 10, 2, 2), 23, 10)
        List1.AddItem TXT$
      End If
    Next VENDE%
    Check4.Value = 1
    '
    Label1 = "Por Facturas CANCELADAS"
    '
    If List1.ListCount > 0 Then
      List1.ListIndex = 0
    End If
    List1.Refresh
    Command3.Enabled = True
    '
End Sub
'
Sub BORRACOMIS()
    '
    MESAN = ""
    '
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) <> "" Then
      PERIO$ = TRIM$(MESANO$(MESASEL$, DES%, HAS%))
    End If
    '
    If PERIO$ = "" Then Exit Sub
    '
    MESAN = PERIO$
    PERIOX$ = TRIM$(PERIO$)
    Call BLANARCH("LICOMACC/NOERR")
    '
    If EXISTE%(LUDAT$ + "CUECOMIS.RFS") Then
       Screen.MousePointer = 11
       Call BLOQARCH("CUECOMIS")
       FIN& = ULTREG&("CUECOMIS")
       JJ& = 0
       For U& = 1 To FIN&
           XX$ = REGLEIDO$("CUECOMIS", U&)
           TIMX% = Asc(Mid$(XX$, 4, 1))
           PER$ = TRIM$(Mid$(XX$, 97, 16))
           ' REGENERA LA CUENTA CORRIENTE DE VENDEDORES (CUECOMIS)
           ' EXCLUYENDO LAS LIQUIDACIONES REALIZADAS EN EL MISMO
           ' PERIODO ELEGIDO (PERIOX$)
           If PER$ <> PERIOX$ Or TIMX% = 2 Then
              JJ& = JJ& + 1
              Call GRAREG("CUECOMIS", XX$, JJ&)
           End If
       Next U&
       Call SETULTREG("CUECOMIS", JJ&)
       Call CIERRARCH("CUECOMIS")
       Screen.MousePointer = 1
    End If
    '\\\OT-007-0076-RG-FIN///
    '
    Screen.MousePointer = 11
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE PerLiComis = '" & PERIO$ & "' "
    '
    'Call APERBASDAT("CABAN")
    'Set AVANCETMP = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
    Dim AVANCETMP As ADODB.Recordset
    Set AVANCETMP = New ADODB.Recordset
    AVANCETMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic

    With AVANCETMP
       On Error Resume Next
       .MoveFirst
       If Err Then
          On Error GoTo 0
          GoTo 90
       End If
       '
       Do While Not .EOF
          '.Edit
          !PerLiComis = ""
          !ImpoComis = 0
          !VendeComis = 0
          .Update
49     .MoveNext
       Loop
    End With
    '
90  List1.Clear
    Screen.MousePointer = 1
    Call COMUNI("Proceso Completo")
    '
End Sub

Sub PRILICOMIS()
    '
    MODOCAL% = 0
    If Option1.Value = True Then MODOCAL% = 1
    If Option3.Value = True Then MODOCAL% = 2
    If Option4.Value = True Then MODOCAL% = 3   ' POR FACTURACIÓN
    If Option5.Value = True Then MODOCAL% = 4   ' POR COBRANZA
    '
    If MODOCAL% = 1 Then
       Call PRILICOMSIL
       Exit Sub
    End If
    '
    If List1.ListIndex >= 0 Then
        TXT$ = List1.TEXT
        VENDE% = Val(Left$(TXT$, 3))
        NOVEN$ = "Clientes de Direccion"
        If VENDE% > 0 Then
            NOVEN$ = ECOARCH$("VENDEDOR", Chr$(VENDE%))
        End If
        '
        CORRESP$ = MESAN.TEXT
        '
        Call COPYSTRU("LICOMIN", "LICOMIV")
        Call BLANARCH("LICOMIV")
        Call HEADERS("LICOMIV", "")
        Call HEADERS("LICOMIV", "Corresponde a: " + CORRESP$)
        Call HEADERS("LICOMIV", "Modalidad: " + Label1)
        Call HEADERS("LICOMIV", "Vendedor: " + NOVEN$)
        '
        FIN& = ULTREG&("LICOMIN")
        JJ& = 0
        '
        For IL& = 1 To FIN&
           Call TRACE(20, IL&, FIN&)
           x$ = REGLEIDO$("LICOMIN", IL&)
           If CVI(Mid$(x$, 19, 2)) = VENDE% Then
              JJ& = JJ& + 1
              Call GRAREG("LICOMIV", x$, JJ&)
           End If
        Next IL&
        '
        Call SETULTREG("LICOMIV", JJ&)
        Call CIERRARCH("*.*")
        Call CONECRUN("*LICOMIV", "Listado de Comisiones Liquidadas")
        '
      Else
        '
        Call MENSERR(24, "Falta Seleccionar Vendedor")
        '
    End If
    '
End Sub
'
Sub PRILICOMSIL()
    '
    If List1.ListIndex >= 0 Then
        TXT$ = List1.TEXT
        VENDE% = Val(Left$(TXT$, 3))
        NOVEN$ = "Clientes de Direccion"
        If VENDE% > 0 Then
            NOVEN$ = ECOARCH$("VENDEDOR", Chr$(VENDE%))
        End If
        '
        CORRESP$ = MESAN.TEXT
        '
        Call COPYSTRU("LICOMACC", "LICOMADD")
        Call BLANARCH("LICOMADD")
        Call HEADERS("LICOMADD", "")
        Call HEADERS("LICOMADD", "Corresponde a: " + CORRESP$)
        Call HEADERS("LICOMADD", "Modalidad: " + Label1)
        Call HEADERS("LICOMADD", "Vendedor: " + NOVEN$)
        '
        FIN& = ULTREG&("LICOMACC")
        JJ& = 0
        '
        For IL& = 1 To FIN&
           Call TRACE(20, IL&, FIN&)
           x$ = REGLEIDO$("LICOMACC", IL&)
           If CVI(Mid$(x$, 33, 2)) = VENDE% Then
              JJ& = JJ& + 1
              Call GRAREG("LICOMADD", x$, JJ&)
           End If
        Next IL&
        '
        Call SETULTREG("LICOMADD", JJ&)
        Call CIERRARCH("*.*")
        Call CONECRUN("*LICOMADD", "Listado de Comisiones Liquidadas")
        '
      Else
        '
        Call MENSERR(24, "Falta Seleccionar Vendedor")
        '
    End If

End Sub
Sub SALCOMI()
  '
  '\\\OT-007-0076-RG-28/02/07///
  Dim SALVEN#(256)
  '
  FIN& = ULTREG&("CUECOMIS")
  Z$ = REGBLAN$("SALCOMIS")
  j& = 0
  For U& = 1 To FIN&
      XX$ = REGLEIDO$("CUECOMIS", U&)
      NVEND% = Asc(Mid$(XX$, 1, 1))
      IMPCOM = CVD(Mid$(XX$, 41, 8))
      IMPABO = CVD(Mid$(XX$, 49, 8))
      SALVEN#(NVEND%) = SALVEN#(NVEND%) + (IMPCOM - IMPABO)
      '
  Next U&
  '
  For y& = 0 To 255
   If Abs(SALVEN#(y&)) > 0.005 Then
      Call REPLA(Z$, Chr$(y&), 1, 1)
      Call REPLA(Z$, MKD$(SALVEN#(y&)), 2, 8)
      j& = j& + 1
      Call GRAREG("SALCOMIS", Z$, j&)
   End If
  Next y&
  '
  Call SETULTREG("SALCOMIS", j&)
  Call CIERRARCH("SALCOMIS")
  '
  Call FINAL("*SALCOMIS")
  '\\\OT-007-0076-RG-FIN///
  '
End Sub
'
Sub LICUECVEN()
'\\\OT-007-0393-FL-19/10/07///
    CUECOPY% = COPYFILE(LUDAT$ + "CUECOMIS.DAT", LUDAT$ + "CUECOMIM.DAT")
    Call FILESORT("CUECOMIM", "", 2, 2, "ASC")
    Call FILESORT("CUECOMIM", "", 1, 1, "ASC")
    
    FIN& = ULTREG&("CUECOMIM")
    
    VACU# = 0: ULTVEN% = 0
    For T& = 1 To FIN&
      Call TRACE(20, T&, FIN&)
      x$ = REGLEIDO$("CUECOMIM", T&)
      VEND% = Asc(Mid$(x$, 1, 1))
      If VEND% <> ULTVEN% Then
        VACU# = 0
        ULTVEN% = VEND%
      End If
      VACU# = VACU# + CVD(Mid$(x$, 41, 8)) - CVD(Mid$(x$, 49, 8))
      Call REPLA(x$, MKD$(VACU#), 57, 8)
      Call GRAREG("CUECOMIM", x$, T&)
    Next T&
    '
    Call DESHASFECHA(DES%, HAS%, PERIO$)
    'Call BLANARCH("CUECVEFE")
    JJ& = 0
    If PERIO$ = "" Then
        Exit Sub
      Else
        FIN& = ULTREG&("CUECOMIM")
        For U& = 1 To FIN&
          Call TRACE(20, U&, FIN&)
          x$ = REGLEIDO$("CUECOMIM", U&)
          '
          FF% = CVI(Mid$(x$, 2, 2))
          If FF% >= DES% Then
            If FF% <= HAS% Then
               JJ& = JJ& + 1
               Call GRAREG("CUECVEFE", x$, JJ&)
            End If
          End If
        Next U&
        Call SETULTREG("CUECVEFE", JJ&)
        Call FINAL("*CUECVEFE")
    End If
    '\\\OT-007-0393-FL-FIN///
End Sub
