VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form OPLIPRE07 
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Presupuestos y Cotizaciones - Opciones"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4815
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   330
      Index           =   0
      Left            =   1470
      ScaleHeight     =   270
      ScaleWidth      =   1110
      TabIndex        =   15
      Top             =   525
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2D3C0&
      Caption         =   "CONTROL"
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
      Left            =   120
      TabIndex        =   9
      Top             =   105
      Width           =   3690
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   645
         Index           =   2
         Left            =   735
         ScaleHeight     =   615
         ScaleWidth      =   2085
         TabIndex        =   10
         Top             =   420
         Width           =   2115
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
            Picture         =   "OPLIPRE07.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Visualizar y Re-Imprimir Presupuestos"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Marcar Como ""No Vigente"""
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
            TabIndex        =   12
            Top             =   105
            Width           =   1425
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2D3C0&
      Caption         =   "CONSULTAS Y LISTADOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   3690
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   645
         Index           =   1
         Left            =   840
         ScaleHeight     =   615
         ScaleWidth      =   2400
         TabIndex        =   4
         Top             =   840
         Width           =   2430
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
            Height          =   630
            Left            =   0
            Picture         =   "OPLIPRE07.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Listado Presupuestos Emitidos por Rango de Fechas"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Listados por Rango de Fechas"
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
            TabIndex        =   6
            Top             =   105
            Width           =   1395
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   645
         Index           =   0
         Left            =   420
         ScaleHeight     =   615
         ScaleWidth      =   2400
         TabIndex        =   1
         Top             =   420
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
            Height          =   630
            Left            =   1850
            Picture         =   "OPLIPRE07.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Visualizar y Re-Imprimir Facturas Proformas"
            Top             =   0
            Width           =   560
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
            Height          =   630
            Left            =   0
            Picture         =   "OPLIPRE07.frx":0C7E
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Visualizar y Re-Imprimir Presupuestos"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Consultas y Re-Impresion"
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
            Left            =   600
            TabIndex        =   3
            Top             =   0
            Width           =   1185
         End
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00624E28&
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   3960
      ScaleHeight     =   3825
      ScaleWidth      =   1560
      TabIndex        =   7
      Top             =   0
      Width           =   1590
      Begin VB.CommandButton Command4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   120
         Picture         =   "OPLIPRE07.frx":12E8
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Recibos de Cobranzas"
         Top             =   960
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton AGenRep 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   105
         Picture         =   "OPLIPRE07.frx":15F2
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Acceso a Consultas Pre-Configuradas"
         Top             =   1680
         Width           =   650
      End
      Begin VB.CommandButton Command2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "OPLIPRE07.frx":18FC
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   200
         Width           =   650
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "OPLIPRE07.frx":1A46
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "OPLIPRE07.frx":1C7A
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "OPLIPRE07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim BRTEMP As Recordset

Sub GENINDIPRES(Optional INIPROFO&)
   '
   Call ABREPRESVEN
   ULTINUM& = XVALO(VALOBADA("PRESUVEN", "NroPres", "NroPres > 0 ORDER BY NroPres DESC"))
   '
   UNUMINDI& = 0
   URE& = ULTREG&("INDIPRES")
   If URE& > 0 Then
      XX$ = REGLEIDO$("INDIPRES", URE&)
      UNUMINDI& = CVS(Left$(XX$, 4))
   End If
   '
   If UNUMINDI& <> ULTINUM& Then
      Screen.MousePointer = 11
      SQLX$ = " SELECT * FROM PRESUVEN WITH(NOLOCK)"
      SQLX$ = SQLX$ + " ORDER BY NroPres ASC "
      '
      Set LISPRETEMP = New ADODB.Recordset
      'LISPRETEMP.CursorLocation = adUseClient
      Set LISPRETEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))    ', FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
      
      With LISPRETEMP
         On Error Resume Next
         .MoveFirst
         If Err < 1 Then
           jj& = 0: NPXA& = 0: FIN& = ULTINUM&
           REBLA$ = REGBLAN("INDIPRES")
           Do While (.EOF = False)
             NPX& = !NroPres
             NCLIE% = !NroClie
             CIXI% = !CodiInt
             FEX = CVINT(!FECHEMIS)
             XXX$ = REBLA$
             '
             '\\\OT-06-0168-WAR-09/05/06///
             CN1$ = "": MARNOVI$ = ""
             On Error Resume Next
             CR1$ = !Cod_Rubro
             NCR1 = !NumeCRubro
             CA1$ = !Cod_Anio
             '\\\OT-06-0497-RG-21/11/06///
             MARNOVI$ = !MarNoVig
             '\\\OT-06-0497-RG-FIN///
             On Error GoTo 0
             '
             If TRIM$(CR1$) <> "" And XVALO(NCR1) > 0 Then
               NCR2$ = TRIM$(Str$(NCR1))
               While Len(NCR2$) < 4: NCR2$ = "0" + NCR2$: Wend
               CN1$ = CR1$ + "-" + NCR2$ + "-" + CA1$ + "   "
             End If
             '\\\OT-06-0168-WAR-FIN///
                NomCli$ = TRIM$(rasocli$(NCLIE%))
                If NCLIE% < 1 Then 'RASON SOCIAL DEL PROSPECT
                    NUMCLI$ = VALOBADA("PRESUVEN", "CODI_CLI", "NroPres = " & NPX& & "")
                    NomCli$ = VALOBADA("BSEGCOM", "RASO_CLI", "CODI_CLI = '" & NUMCLI$ & "'")
                End If

             YY2$ = TRIM$(AJUSTI$(CN1$ + NomCli$, 30)) + " - "
             If NPX& <> NPXA& Then
                  YY2$ = YY2$ + TRIM$(CODEXT$(CIXI%))
                Else
                  YY2$ = YY2$ + "VARIOS"
             End If
             YY2$ = YY2$ + " - " + FECHATEX$(FEX)
             '\\\OT-06-0497-RG-21/11/06///
             If MARNOVI$ <> "" Then
                YY2$ = "(N/Vig) " + YY2$
             End If
             '\\\OT-06-0497-RG-FIN///
             '
             Call REPLA(XXX$, MKS$(NPX&), 1, 4)
             Call REPLA(XXX$, YY2$, 5, 60)
             If NPX& <> NPXA& Then
               jj& = jj& + 1
               Call TRACE(20, jj&, FIN&)
             End If
             Call GRAREG("INDIPRES", XXX$, jj&)
             '
             NPXA& = NPX&
           .MoveNext
           Loop
         End If
      End With
      LISPRETEMP.Close
      Set LISPRETEMP = Nothing
      '
      Call SETULTREG("INDIPRES", jj&)
      Call CIERRARCH("INDIPRES")
      '
   End If
   '
   Call FRESHECHO("INDIPRES")
   Screen.MousePointer = 1
   
End Sub
'
Sub ACTINDIPRES(NPX&)
   '
   Call ABREPRESVEN
   SQLX$ = " SELECT * FROM PRESUVEN WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE NroPres = " & NPX&
   '
   Dim LISPRETEMP As ADODB.Recordset
   Set LISPRETEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))    ', FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
   '
   With LISPRETEMP
      If Not (.BOF And .EOF) Then
        XXX$ = REGBLAN("INDIPRES")
        jj& = 0: NPXA& = 0: YY2$ = ""
        Do While (.EOF = False)
           NCLIEX = !NroClie
           CIXI% = !CodiInt
           FEX = CVINT(!FECHEMIS)
           '
           CN1$ = "": MARNOVI$ = ""
           On Error Resume Next
           CR1$ = !Cod_Rubro
           NCR1 = !NumeCRubro
           CA1$ = !Cod_Anio
           '\\\OT-06-0497-RG-21/11/06///
           MARNOVI$ = !MarNoVig
           '\\\OT-06-0497-RG-FIN///
           On Error GoTo 0
           '
           If TRIM$(CR1$) <> "" And XVALO(NCR1) > 0 Then
             NCR2$ = TRIM$(Str$(NCR1))
             While Len(NCR2$) < 4: NCR2$ = "0" + NCR2$: Wend
             CN1$ = CR1$ + "-" + NCR2$ + "-" + CA1$ + "   "
           End If
           '\\\OT-06-0168-WAR-FIN///
           NomCli$ = TRIM$(rasocli$(NCLIEX))
           If NCLIEX < 1 Then 'RASON SOCIAL DEL PROSPECT
               CDXCLI$ = VALOBADA("PRESUVEN", "CODI_CLI", "NroPres = " & NPX& & "")
               NomCli$ = VALOBADA("BSEGCOM", "RASO_CLI", "CODI_CLI = '" & CDXCLI$ & "'")
           End If

           YY2$ = TRIM$(AJUSTI$(CN1$ + NomCli$, 30)) + " - "
           If NPX& <> NPXA& Then
                YY2$ = YY2$ + TRIM$(CODEXT$(CIXI%))
              Else
                YY2$ = YY2$ + "VARIOS"
           End If
           YY2$ = YY2$ + " - " + FECHATEX$(FEX)
           '\\\OT-06-0497-RG-21/11/06///
           If MARNOVI$ <> "" Then
              YY2$ = "(N/Vig) " + YY2$
           End If
           NPXA& = NPX&
        .MoveNext
        Loop
        '
        If YY2$ <> "" Then
           Call REPLA(XXX$, MKS$(NPX&), 1, 4)
           Call REPLA(XXX$, YY2$, 5, 60)
           Call FILTERPUTRECORD("INDIPRES", 1, MKS$(NPX&), XXX$)
        End If
        Call CIERRARCH("INDIPRES")
      End If
   End With
   '
End Sub

Sub AGenRep_Click()
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

'
Sub Command1_Click()
    Command1.Enabled = False
    Screen.MousePointer = 11
    'NUEVO
    Call GENINDIPRES
    '
3   Screen.MousePointer = 1
    Call SELECHO("INDIPRES")
    NPX& = XVALO(VALACT1$("INDIPRES"))
    If NPX& < 1 Then GoTo 99
    '
'    TIDOCUM$ = "Presup"
'    NUDOCU$ = " " + TRIM$(Str$(NPX&)) + " -"
    Screen.MousePointer = 11
    '
    'BUSCA EL DOCUMENTO
    DOCUNU& = 0
    Call MUESTRADOC(DOCUNU&, "", "PRESUPUESTO", Str$(NPX&), Str$(NPX&))
    GoTo 3
    '
    
'''    '
'''    For UI& = ULTREG&("PDOCLST") To 1 Step -1
'''      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
'''      If InStr(XX$, TIDOCUM$) > 4 Then
'''        PPXX% = InStr(XX$, "Nro.")
'''        If PPXX% > 0 Then
'''          XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
'''        End If
'''        If InStr(XX$, NUDOCU$) > 4 Then
'''          DOCUNU& = CVS(Left$(XX$, 4))
'''          GoTo 5
'''        End If
'''      End If
'''    Next UI&
'''    Screen.MousePointer = 1
'''    Call COMUNI("Documento no Encontrado")
'''    GoTo 99
'''    '
'''5   Call MUESTRADOC(DOCUNU&)
'''    GoTo 3
    
'''    For KKI% = 0 To 5
'''      EDITFORM.Picture1(KKI%).Cls
'''      EDITFORM.Picture1(KKI%).DrawWidth = 3
'''      EDITFORM.Picture1(KKI%).DrawStyle = 0
'''      EDITFORM.Picture1(KKI%).FillStyle = 1
'''    Next KKI%
'''    '
'''    REDOCU$ = REGACT$("PDOCLST")
'''    URE& = ULTREG&("PDOCSPL")
'''    LI& = 0: LS& = URE&
'''6   E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
'''    L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
'''    DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'''    If DCC& < DOCUNU& Then LI& = L&: GoTo 6
'''    If DCC& > DOCUNU& Then LS& = L&: GoTo 6
'''    GoTo 8
'''    '
'''7   Screen.MousePointer = 1
'''    Call MENSERR(24, "Documento no Registrado")
'''    GoTo 99
'''    '
'''8   TPSP$ = ""
'''    Screen.MousePointer = 11
'''    While L& > 1
'''      DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'''      If DCC& < DOCUNU& Then GoTo 9
'''      L& = L& - 1
'''    Wend
'''    '
'''9   CAPAGINAS% = 0
'''    For KKL& = L& To URE&
'''      TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
'''      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
'''      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
'''        TPSP$ = TPSP$ + Mid$(TBUF$, 5)
'''      End If
'''    Next KKL&
'''    '
'''95  SPOOLSTRING$ = TPSP$
'''    For KKI% = 0 To 5
'''      EDITFORM.Picture1(KKI%).Cls
'''      EDITFORM.Picture1(KKI%).Top = 0
'''      EDITFORM.Picture1(KKI%).Refresh
'''      EDITFORM.Picture1(KKI%).Height = 7170
'''      EDITFORM.Picture1(KKI%).Width = 11700
'''      EDITFORM.Picture1(KKI%).DrawWidth = 3
'''      EDITFORM.Picture1(KKI%).DrawStyle = 0
'''      EDITFORM.Picture1(KKI%).FillStyle = 1
'''    Next KKI%
'''    ALTUPAGINA = 7170
'''    TOPEPAGINA = 0
'''    PAGINACTIVA% = 0
'''    CAPAGINAS% = 0
'''    '
'''10  PPXX% = InStr(TPSP$, Chr$(255))
'''    If PPXX% > 0 Then
'''      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
'''      TPSP$ = Mid$(TPSP$, PPXX% + 1)
'''      Call TRACE(20, FIN& - Len(TPSP$), FIN&)
'''      '
'''      If Len(COPRIAT$) > 0 Then
'''        COPRI% = Asc(Left$(COPRIAT$, 1))
'''        ATRIP$ = Mid$(COPRIAT$, 2)
'''        Call SHOWCOMMAND(COPRI%, ATRIP$)
'''      End If
'''      GoTo 10
'''      '
'''    End If
'''    '
'''    UI& = UI& + 1
'''    XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
'''    If InStr(XX$, TIDOCUM$) > 4 Then
'''      If InStr(XX$, NUDOCU$) > 4 Then
'''        DOCUNU& = CVS(Left$(XX$, 4))
'''        GoTo 7
'''      End If
'''    End If
'''    '
'''    Screen.MousePointer = 1
'''    PAGINACTIVA% = 0
'''    For KKI% = 0 To 5
'''       EDITFORM.Picture1(KKI%).Visible = False
'''    Next KKI%
'''    EDITFORM.Picture1(PAGINACTIVA%).Visible = True
'''    EDITFORM.Picture1(PAGINACTIVA%).Refresh
'''    EDITFORM.VScroll1.Min = 0
'''    EDITFORM.VScroll1.Value = 0
'''    EDITFORM.VScroll1.Max = 1
'''    If ALTUPAGINA > EDITFORM.Frame1.Height Then
'''       EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
'''    End If
'''    EDITFORM.Command3.Enabled = False
'''    EDITFORM.Show 1
'''    '
'''    GoTo 3
    '
99  Command1.Enabled = True
    Screen.MousePointer = 1
    '
End Sub

Private Sub Command2_Click()
    '
    Call CIERRARCH("*.*")
    Hide
    '
End Sub

Sub Command3_Click()
    '
    Command3.Enabled = False
2   Call OPLIPRE07.GENINDIPRES
    Call FRESHECHO("INDIPRES")
    '
3   Screen.MousePointer = 1
    Call SELECHO("INDIPRES")
    NPX& = XVALO(VALACT1$("INDIPRES"))
    If NPX& < 1 Then GoTo 99
    '\\\OT-06-0497-RG-21/11/06///
    RESPU% = COMALTER%("Confirma Marcar Presupuesto Nro." + Str$(NPX) + " como 'No Vigente' \\ No, Consevar \ Si, Marcar")
    If RESPU% <> 2 Then GoTo 99
    '
    Call ABREPRESVEN
    'Jandy SQL
    strSQL = "Update Presuven set MarNoVig = '*'"
    strSQL = strSQL + " where NROPRES = " & NPX&
     FLEXCONN.Execute (FILTSQL$(strSQL))
    '
    Call ACTINDIPRES(NPX&)
    '
    Call COMUNI("Transacción Completada")
    GoTo 2
    '
99  Command3.Enabled = True
    Screen.MousePointer = 1
    '
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    '
    'MUESTRO SELECCION DE CLIENTES
10  Call SELECHO("DEUDOR1")
    NC% = XVALO(VALACT1$("DEUDOR1"))
    If NC% < 1 Then
        GoTo 99
    End If
    '
    'VALIDO QUE POSEA RECIBOS SI NO PRESENTO MENSAJE
    CONDI$ = "TipoMovim= 'RCOB' AND NuClien= " & NC% & " "
    If CantRegistros("CAJABANC", CONDI$) < 1 Then
       Call COMUNI("No Se Encontraron Recibos para el Cliente Seleccionado")
       GoTo 10
    End If
    '
    'MUESTRO LOS RECIBOS SEGUN CLIENTE
    CONDI$ = CONDI$ + " ORDER BY FECHEMISI,CODICOM_LAR"
30  Call CARGALISEL("LIRECIBO", "CAJABANC", "FECHEMISI/D8;CODICOM_LAR/A18", CONDI$, "NOMESS")
    Call SELECHO("LIRECIBO")
    '
    REXIBO$ = TRIM$(VALACT2$("LIRECIBO"))
    If REXIBO$ = "" Then GoTo 10
    '
    'PRESENTO EL STDSCREEN
50  FILTX$ = REXIBO$
    Call STDFORM("RECI-GBL", FILTX$)
    GoTo 30

    
99  Command4.Enabled = True
End Sub




Sub Command5_Click()
    Command5.Enabled = False
    Screen.MousePointer = 11
    'NUEVO
    INICIPRO& = XVALO(Control("PRESUPUE", "INIPROFO"))
    If INICIPRO& < 1 Then
       Call COMUNI("Atención !! No se ha Inicializado la Numeración de Facturas Proforma\Se Mostrarán Todos Los Presupuestos Existentes\Controle La Información de los Mismos")
    End If
    CONDI$ = "NROPRES >= " & INICIPRO&
    
    Call CARGALISEL("PREPROFO", "PRESUVEN", "NROPRES/F7; NroClie/F6+DEUDOR12/A32;FechEmis/D10;MarNoVig/A6", CONDI$, "NOMESS/DISTINCT")
'    Call GENINDIPRES(INICIPRO&)
    '
3   Screen.MousePointer = 1
    Call SELECHO("PREPROFO/LISTADO DE FACTURAS PROFORMA                                     (MARCADOS CON ' * ' - NO VIGENTES)")
    NPX& = XVALO(VALACT1$("PREPROFO"))
    If NPX& < 1 Then GoTo 99
    '
    Screen.MousePointer = 11
    '
    ITEMUNICO% = 0
    CONDI$ = "NROPRES = " & NPX&
    NCX = XVALO(VALOBADA("PRESUVEN", "NROCLIE", CONDI$, "NOMESS"))
    DESCRIPX$ = UCase$(TRIM$(VALOBADA("PRESUVEN", "ESPECIF", CONDI$, "NOMESS")))
    CI1% = XVALO(VALOBADA("PRESUVEN", "CodiInt", CONDI$, "NOMESS"))
    '
    'COMPARO ESPECIFICACION CON LA DESCRIPCION DEL ARTICULO CUANDO ES IGUAL ES POR QUE ES LISTA DE REPUESTOS, CASO CONTRARIO ES ITEM UNICO
    If DESCRIPX$ <> UCase$(TRIM$(DESCRIT0$(CI1%))) Then ITEMUNICO% = 1
    FILTX$ = TRIM$(NPX&) & ";" & TRIM$(Str$(NCX))
    If ITEMUNICO% > 0 Then
       FRMITEMUNICO$ = Control("PRESIUNI", "PROFLREP")
       Call STDFORM(FRMITEMUNICO$, FILTX$)
    Else
       FRMLISTAREP$ = Control("PRESUPUE", "PROFLREP")
       Call STDFORM(FRMLISTAREP$, FILTX$)
    End If
    GoTo 3
99 Command5.Enabled = True
End Sub

 Sub Command6_Click()
    Command6.Enabled = False
    
10  FEX = HOY(HO$)
    Has% = FEX
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    Des% = FECHANUM(DESS$)
    VDEF$ = MKI$(Des%) + MKI$(Has%)
    '
    OBX$ = OBJPLA$("DESHASFECHCLIN", VDEF$) '\\\OT-06-0268-WAR-29/06/06///
    If OBX$ = "" Then
        GoTo 99
    End If
    '
    Des% = CVI(Left$(OBX$, 2))
    FEX = Des%: DESX$ = FECHATEX$(FEX)
    Has% = CVI(Mid$(OBX$, 3, 2))
    FEX = Has%: HASS$ = FECHATEX$(FEX)
    If Has% < Des% Then
      Call MENSERR(24, "Rango de Fechas no Válido")
      GoTo 10
    End If
    NCLIE1 = CVS(Mid$(OBX$, 5, 4)) '\\\OT-06-0268-WAR-29/06/06///
    '
    Screen.MousePointer = 11
    'Call COPYSTRU("PRESUPU", "LIPRESU")
    
    'Nuevo
    Call ABREPRESVEN
    'jandy
    Desde = FETEXCOR1$(Des%)
    Hasta = FETEXCOR1$(Has%)
    '
    SQLX$ = " SELECT * FROM PresuVen "
    'SQLX$ = SQLX$ + " WHERE cdbl(FechEmis) >=   " & CDbl(CVDAT(DES%)) & "  "
    'SQLX$ = SQLX$ + " AND cdbl(FechEmis) <=  " & CDbl(CVDAT(HAS%)) & "   "
    SQLX$ = SQLX$ + " WHERE FechEmis >= '" & Desde & "'"
    SQLX$ = SQLX$ + " AND FechEmis <= '" & Hasta & "'"
    If NCLIE1 > 0 Then
      SQLX$ = SQLX$ + " AND NroClie = " & NCLIE1 & " "
    End If
    SQLX$ = SQLX$ + " ORDER BY NroPres "
    '
    'Set PRESUTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    Dim PRESUTEMP As ADODB.Recordset
    Set PRESUTEMP = New ADODB.Recordset
    PRESUTEMP.CursorLocation = adUseClient
    PRESUTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
   
    With PRESUTEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        jj& = 0
        X$ = REGBLAN("LIPRESU")
        Do While (.EOF = False)
          On Error Resume Next
          Call REPLA(X$, MKS$(!NroPres), 1, 4)
          Call REPLA(X$, MKI$(CVINT(!FECHEMIS)), 5, 2)
          Call REPLA(X$, MKS$(!NroClie), 7, 4)
            MONEMISI% = !MonedEmi
            If MONEMISI% < 1 Then MONEMISI% = 1
          Call REPLA(X$, Chr$(!ConPag), 68, 1)
          Call REPLA(X$, Chr$(!VENDED), 69, 1)
          Call REPLA(X$, Chr$(MONEMISI%), 70, 1)
          Call REPLA(X$, MKI$(!CodiInt), 71, 2)
          Call REPLA(X$, MKS$(!PreList), 73, 4)
          Call REPLA(X$, !TDESCUEN, 77, 14)
          Call REPLA(X$, MKS$(!PREUNIT), 99, 4)
          Call REPLA(X$, MKS$(!Cantit), 103, 4)
            ITOPES# = !PREUNIT * !Cantit * TICAMONE(MONEMISI%)
          Call REPLA(X$, MKD$(ITOPES#), 107, 8)
          Call REPLA(X$, MKI$(CVINT(!FechVent)), 115, 2)
          Call REPLA(X$, MKS$(!NroPedi), 117, 4)
          On Error GoTo 0
          jj& = jj& + 1
          Call GRAREG("LIPRESU", X$, jj&)
          .MoveNext
        Loop
      End If
      On Error GoTo 0
      Call SETULTREG("LIPRESU", jj&)
      Call CIERRARCH("LIPRESU")
    End With
    PRESUTEMP.Close
    Set PRESUTEMP = Nothing
    '
    Call HEADERS("LIPRESU", "")
    Call HEADERS("LIPRESU", "Periodo: " + DESS$ + " - " + HASS$)
    '
    Call CONECRUN("*LIPRESU", "Listado de Presupuestos")
    Screen.MousePointer = 1
    '
99  Command6.Enabled = True
    '
End Sub

Sub RECODOSI(CIXI%, CIXJ%, NPX&)
   '
   CIBIN$ = MKI$(CIXI%)
   CIXJ% = CIXI%
   If Asc(Mid$(CIBIN$, 2, 1)) = 0 Then
     If NPX& > 1600 Then
       If NPX& <= 1614 Then
         If NPX& = 1601 Then CIXJ% = CODINT%("041130M-PVT0")
         If NPX& = 1602 Then CIXJ% = CODINT%("042121T-PHH0")
         If NPX& = 1603 Then CIXJ% = CODINT%("041015M-PVT0")
         If NPX& = 1604 Then CIXJ% = CODINT%("041015M-PVT0")
         If NPX& = 1605 Then CIXJ% = CODINT%("041130M-PVT0")
         If NPX& = 1606 Then CIXJ% = CODINT%("041130M-PVT0")
         If NPX& = 1607 Then CIXJ% = CODINT%("042501T-PHH0")
         If NPX& = 1608 Then CIXJ% = CODINT%("44234")
         If NPX& = 1609 Then CIXJ% = CODINT%("44234")
         If NPX& = 1610 Then CIXJ% = CODINT%("LD54S2-PHCA-D13")
         If NPX& = 1611 Then CIXJ% = CODINT%("044501T-PTT0")
         If NPX& = 1612 Then CIXJ% = CODINT%("044251T-PTT0")
         If NPX& = 1613 Then CIXJ% = CODINT%("041130M-PVT0")
         If NPX& = 1614 Then CIXJ% = CODINT%("A4190")
       End If
     End If
   End If
   If CIXJ% < 1 Then CIXJ% = CIXI%
   '
End Sub


Private Sub Form_Load()
    '
    'BOTON DE RECIBOS POR FRM STDSCREEN
    HOII% = HOY(HO$)
    DIAX% = FECHANUM("30/07/10")
    If HOII% <= DIAX% Then Command4.Visible = True
    '
    Call APLICACIONSKINS(Me, Picture2)
End Sub
