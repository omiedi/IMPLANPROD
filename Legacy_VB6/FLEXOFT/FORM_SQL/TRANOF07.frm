VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form TRANOF07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selección de O-Fabricación a Transferir"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   7635
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   20
      Top             =   720
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.ListBox LISNUM 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3120
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   0
      ToolTipText     =   "Tilde para seleccionar"
      Top             =   600
      Width           =   6465
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00E0E0E0&
      Height          =   435
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   6405
      TabIndex        =   7
      Top             =   120
      Width           =   6465
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "O-Fabricacion  Codigo - Descripción"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   225
         Left            =   120
         TabIndex        =   8
         Top             =   105
         Width           =   10515
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00400000&
      Height          =   5775
      Left            =   6720
      ScaleHeight     =   5715
      ScaleWidth      =   1455
      TabIndex        =   3
      Top             =   0
      Width           =   1515
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
         Height          =   640
         Left            =   120
         Picture         =   "TRANOF07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   120
         Width           =   650
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
         Height          =   650
         Left            =   120
         Picture         =   "TRANOF07.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Refresca - Invierte Selección"
         Top             =   2370
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   4
         Top             =   3945
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   5
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton BOTREC 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   120
         Picture         =   "TRANOF07.frx":0294
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Aprueba la Selección"
         Top             =   3075
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "TRANOF07.frx":059E
         Top             =   4290
         Width           =   1515
      End
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   10200
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   5400
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Depósitos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   11
      Top             =   3720
      Width           =   6475
      Begin VB.Label Label7 
         Height          =   255
         Left            =   720
         TabIndex        =   19
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label6 
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dep.Consumo: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4560
         TabIndex        =   17
         Top             =   300
         Width           =   1320
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dep.Entrada: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   960
         TabIndex        =   16
         Top             =   300
         Width           =   1800
      End
      Begin VB.Label Label8 
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
         Left            =   240
         TabIndex        =   15
         Top             =   660
         Width           =   2895
      End
      Begin VB.Label Label1 
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
         Left            =   2760
         TabIndex        =   14
         Top             =   240
         Width           =   315
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
         Left            =   3360
         TabIndex        =   13
         Top             =   660
         Width           =   2895
      End
      Begin VB.Label Label3 
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
         Left            =   5880
         TabIndex        =   12
         Top             =   240
         Width           =   315
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "TRANOF07.frx":24C0
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "TRANOF07.frx":26F4
      TabIndex        =   21
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   5500
      Y1              =   3255
      Y2              =   3255
   End
   Begin VB.Line Line2 
      X1              =   5505
      X2              =   5505
      Y1              =   495
      Y2              =   3245
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   330
      Left            =   7440
      TabIndex        =   9
      Top             =   1140
      Width           =   1695
   End
End
Attribute VB_Name = "TRANOF07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'


Private Sub BOTREC_Click()
    '
    Call CREACAMPO("", "RESERVA", "CANTENT", 6, 0)
    Call PoneEnCeroNull("RESERVA", "CANTENT", REGAF&)
    '
    MAXOF% = 16: NN% = 0
    Dim LISTAOF() As String * 12
    ReDim LISTAOF(MAXOF%)
    '
    Condi1$ = "": CONDI2$ = ""
    For i = 0 To LISNUM.ListCount - 1
        If LISNUM.Selected(i) = True Then
            TXTX$ = LISNUM.List(i)
            OF$ = TRIM$(Mid$(TXTX$, 1, 12))
            NN% = NN% + 1
            If NN% > MAXOF% Then
                MAXOF% = MAXOF% + 1
                ReDim LISTAOF(MAXOF%)
            End If
            LISTAOF(NN%) = OF$
            Condi1$ = Condi1$ & "IDSUBOR='" & OF$ & "' OR "
            CONDI2$ = CONDI2$ & "DOSECCOR='" & OF$ & "' OR "
        End If
    Next i
    '
    If TRIM$(Condi1$) = "" Then Exit Sub
    '
    CONDI$ = Mid$(Condi1$, 1, Len(Condi1$) - 3)
    '
    SQLX$ = "SELECT COD_INTE, SUM(CANTRES) AS CANRES, SUM(CANTENT) AS CANCON FROM RESERVA WITH(NOLOCK) WHERE " & CONDI$ & " GROUP BY COD_INTE"
    Set RESERTMP = READSET(SQLX$)
    '
    jj& = 0: CONTA% = 0
    REBLA$ = REGBLAN("!RESERVA")
    Call SETULTREG("!RESERVA", 0)
    With RESERTMP
      Do While (.EOF = False)
          Y$ = REBLA$
          CII% = XVALO(!cod_inte)
          If ESTRAZABLE(CII%) = 2 Then
              CONTA% = CONTA% + 1 ' valida si hay codigos trazables
            Else
              Call REPLA(Y$, MKI$(!cod_inte), 1, 2)
              Call REPLA(Y$, MKI$(!cod_inte), 15, 2)
              Call REPLA(Y$, MKD$(!CANRES), 19, 8)
              Call REPLA(Y$, MKD$(!CANCON), 27, 8)
              '
              jj& = jj& + 1
              Call GRAREG("!RESERVA", Y$, jj&)
          End If
13    .MoveNext
      Loop
   End With
   RESERTMP.Close
   Set RESERTMP = Nothing
   Call SETULTREG("!RESERVA", jj&)
   Call CIERRARCH("!RESERVA")
   '
   DepOrig% = XVALO(Label1)
   DEPODEST% = XVALO(Label3)
   ARMADOR% = XVALO(Label6)
   '
   CONDI$ = Mid$(CONDI2$, 1, Len(CONDI2$) - 3)
   FIN& = ULTREG&("!RESERVA")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("!RESERVA", U&)
      ORFA$ = Mid$(XX$, 3, 12)
      CI% = CVI(Left$(XX$, 2))
      CANTI = CVD(Mid$(XX$, 19, 8))
      CENTR = CVD(Mid$(XX$, 27, 8))
      CANTI = CANTI - CENTR
      '
      If CANTI < 0 Then CANTI = 0
      If TRIM$(UCase$(Unimed$(CI%))) = "GK" Then CANTI = CANTI / 1000
      '
      STODEST = STODEPOS(CI%, DEPODEST%)
'      CONDI2$ = "COD_INTE=" & CI% & " AND (" & CONDI$ & ")"
'      QENT = XVALO(VALOBADA("MOVISTO", "SUM(CANTSALID)", CONDI2$))
'      QSAL = XVALO(VALOBADA("MOVISTO", "SUM(CANTINGRE)", CONDI2$))
'      STODEST = QENT - QSAL
      '
      STOCENT = STODEPOS(CI%, DepOrig%)
      If STOCENT < 0 Then STOCENT = 0
      NECNETA = CANTI - STODEST
      If NECNETA < 0 Then NECNETA = 0
      '
      VALEMATX = NECNETA ' EL VALE SE PRESENTA POR LA NECESIDAD NETA
      If NECNETA > STOCENT Then ' SI NO HAY STOCK CONSUME HASTA EL STOCK
         VALEMATX = STOCENT
      End If
      '
      If VALEMATX < 0.05 Then VALEMATX = 0
       '
      ZZ$ = REGBLAN$("VALEMAT")
      Call REPLA(ZZ$, MKI$(CI%), 1, 2)
      Call REPLA(ZZ$, MKS$(CANTI), 3, 4) ' NECESIDAD BRUTA
      Call REPLA(ZZ$, MKS$(CENTR), 7, 4)     ' STOCK DEPOSITO DESTINO
      Call REPLA(ZZ$, MKS$(NECNETA), 11, 4)    ' NECESIDAD NETA
      Call REPLA(ZZ$, MKS$(STOCENT), 15, 4)    ' STOCK DEP.CENTRAL
      Call REPLA(ZZ$, MKS$(VALEMATX), 19, 4)   ' VALE A EMITIR
      Call GRAREG("!VALEMAT", ZZ$, U&)
   Next U&
   Call SETULTREG("!VALEMAT", FIN&)
   Call CIERRARCH("!VALEMAT")
   '
25 TITUPA$ = "Vale de TRANSFERENCIA" + " - de " + TRIM$(ECOARCH("TADEPOSI", Chr$(DepOrig%))) + " a " + TRIM$(ECOARCH$("TADEPOSI", Chr$(DEPODEST%)))
   VTB% = VENTABU%("VALEMAT9" + "/TITUPAN=" + TITUPA$)
   If VTB% < 0 Then Exit Sub

   If PUEDEMOVI%(DepOrig%) < 0 Then GoTo 25 ' CANCELA SALIDA
   ' FIN CONTROL DE NEGATIVOS
   '
   If COMALTER%("Confirme la Transferencia a Realizar\\Cancelar\Confirmar Transferencia") <> 2 Then Exit Sub
   
   ' GRABAR TRANSFERENCIA
   '
   CMOV$ = TRIM$(Control$("", "CODVALMA"))
   If CMOV$ = "" Then CMOV$ = "TR"
   '
   NRO& = ProNroComprobante(CMOV$)
   '
   Call BLOQARCH("MOVISTO")
   '
   Screen.MousePointer = 11
   NC% = 0
   VUNI# = 0
   MONE$ = ""
   'LOTE$ = ""
   FECHA% = HOY(HOS$)
   DOPRI$ = "TR." + TRIM$(str$(NRO&))
   'DOSEC$ = "TR." + Mid$(HO$, 8, 1) + Mid$(HO$, 4, 2) + Left$(HO$, 2)
   DOSEC$ = TRIM$(ORFA$) ' NUMERO DE ORDEN DE FABRICACION
    While Len(DOSEC$) < 10
      DOSEC$ = DOSEC$ + " "
    Wend
                              '
   NORI% = 0
   For U& = 1 To ULTREG&("!VALEMAT")
      XX$ = REGLEIDO$("!VALEMAT", U&)
      If CVS(Mid$(XX$, 19, 4)) >= 0.05 Then
         '
         CI1% = CVI(Mid$(XX$, 1, 2))
         CAN = CVS(Mid$(XX$, 19, 4))
         REFE$ = "TR.a " + TRIM$(ECOARCH$("TADEPOSI", Chr$(DepOrig%)))
         '
         CANTI = (-1) * CAN
         NORFI% = NORI% + 1
         '
         LOTE$ = TRIM$(Mid$(XX$, 49, 16)) ' INICIALIZAMOS POR SI NO ENTRA NO QUEDE ANTERIOR
         Call MOVISTOK(FECHA%, CI1%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NORI%, VUNI#, 0, NC%, DepOrig%, CANTI)
         '
         REFE$ = "TR.de " + TRIM$(ECOARCH$("TADEPOSI", Chr$(DEPODEST%)))
         '
         CANTI = CAN
         Call MOVISTOK(FECHA%, CI1%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NORI%, VUNI#, 0, NC%, DEPODEST%, CANTI)
         '
         For II% = 1 To NN%
            If CANTI > 0.005 Then
                OF$ = TRIM$(LISTAOF$(II))
                CONDI$ = "IDSUBOR='" & OF$ & "' AND COD_INTE=" & CI1%
                    CRE = XVALO(VALOBADA("RESERVA", "CANTRES", CONDI$))
                    CEN = XVALO(VALOBADA("RESERVA", "CANTENT", CONDI$))
                SALDO = CRE - CEN
                If SALDO > 0.005 Then
                    valor = SALDO
                    If SALDO > CANTI Then valor = CANTI
                    Call ACTUREGISTRO("RESERVA", "CANTENT", CONDI$, CEN + valor, REGAF&)
                    CANTI = CANTI - valor
                End If
            End If
         Next II%
         '
      End If
   Next U&
   '
   Call LIBARCH("MOVISTO")
   
   Call BAJALDISCO
   '
   'IMPRIME EL VALE DE TRANSFERENCIA
   '
   If ARMADOR% < 1 Then
       '
       CANCELPRINT% = 0
       IMPR$ = TRIM$(Control$(IDENTER$, "SUIMPVAL"))
       If IMPR$ = "SI" Then CANCELPRINT% = 1
       '
60     FORIPRE$ = Control$("", "FORVALMA")
       If FORIPRE$ <> "" Then
          '
          FECHDOC$ = HOS$
          NUMEDOC$ = TRIM$(str$(NRO&))
          While Len(NUMEDOC$) < 6
             NUMEDOC$ = "0" + NUMEDOC$
          Wend
          TIPODOC$ = "Vale de Materiales"
          '
          CODPARAM$(1) = "DEPOSITO": DOCPARAM$(1) = TRIM$(str$(DepOrig%)) + " - " + TRIM$(ECOARCH$("TADEPOSI", Chr$(DepOrig%)))
          CODPARAM$(2) = "DESTINO": DOCPARAM$(2) = TRIM$(str$(DEPODEST%)) + " - " + TRIM$(ECOARCH$("TADEPOSI", Chr$(DEPODEST%)))
          CODPARAM$(3) = "DOCUORIG": DOCPARAM$(3) = TRIM$(NUORFAB$)
          CAPARDOC% = 3
          '
          CODTABU1$(1) = "COD-MAT"
          CODTABU1$(2) = "DES-MAT"
          CODTABU1$(3) = "UNIMED"
          CODTABU1$(4) = "CANTIDAD"
          CODTABU1$(5) = "CAKILOS"
          CODTABU1$(6) = "REF-MAT1"
          CACOLTAB1% = 6
          '
          DESTIDOC% = 0
          CAITAB1% = 0
          CAITCRU% = 0
          '
          For U& = 1 To ULTREG&("!VALEMAT")
             XX$ = REGLEIDO$("!VALEMAT", U&)
             If CVS(Mid$(XX$, 19, 4)) >= 0.05 Then
                '
                CICRU% = CVI(Mid$(XX$, 1, 2))
                CACOL1$ = str$(CVS(Mid$(XX$, 19, 4)))
                KICOL1$ = str$(CVS(Mid$(XX$, 19, 4)) * PESUNI(CICRU%))
                '
                CAITAB1% = CAITAB1% + 1
                TETABU1$(1, CAITAB1%) = CODEXT$(CICRU%)
                TETABU1$(2, CAITAB1%) = DESCRIT0$(CICRU%)
                TETABU1$(3, CAITAB1%) = Unimed$(CICRU%)
                TETABU1$(4, CAITAB1%) = CACOL1$
                TETABU1$(5, CAITAB1%) = KICOL1$
                TETABU1$(6, CAITAB1%) = UBISTO$(CICRU%)
                '
             End If
          Next U&
          '
    
          If CAITAB1% < 1 Then
             Screen.MousePointer = 1
             Exit Sub
          End If
          '
          Call PRINTDOC("FORVALMA")   ' Vale de Materiales
          Call BAJALDISCO
          '
       End If
   Else
        NPROVE% = XVALO(Label7)
        Call OFABRI07.GeneraRemiConsig(NUORFAB$, NPROVE%)
   End If
   '
   Call COMUNI("Transferencia Realizada")
   Screen.MousePointer = 1
   Unload Me
   
End Sub

Private Sub Command1_Click()
   Command1.Enabled = False
   TPI = LISNUM.TopIndex
   For U& = 1 To LISNUM.ListCount
     If LISNUM.Selected(U& - 1) = True Then
         LISNUM.Selected(U& - 1) = False
       Else
         LISNUM.Selected(U& - 1) = True
     End If
   Next U&
   On Error Resume Next
   LISNUM.TopIndex = TPI
   On Error GoTo 0
   DoEvents
   Command1.Enabled = True

End Sub

Private Sub Command2_Click()
   LISNUM.Clear
   Unload Me
End Sub


Private Sub Command3_Click()
   If DERACCE%("CAMBIDEP", "CAMBIO DE DEPOSITOS") > 1 Then
    Call SELECHO("TADEPOSI")
    If VALACT1$("TADEPOSI") <> "" Then
      Label3 = TRIM$(VALACT1$("TADEPOSI"))
    End If
   End If
End Sub

Private Sub Command6_Click()
   If DERACCE%("CAMBIDEP", "CAMBIO DE DEPOSITOS") > 1 Then
    Call SELECHO("TADEPOSI")
    If VALACT1$("TADEPOSI") <> "" Then
      Label1 = TRIM$(VALACT1$("TADEPOSI"))
    End If
   End If
End Sub

Private Sub Form_Load()
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
End Sub

Private Sub label1_Change()
   Label8 = ECOARCH$("TADEPOSI", Chr$(XVALO(Label1)))
End Sub

Private Sub Label3_Change()
   Label2 = ECOARCH$("TADEPOSI", Chr$(XVALO(Label3)))
End Sub


