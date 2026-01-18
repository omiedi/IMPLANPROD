VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form SITUSTO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Situación de Stock"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   12945
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   12945
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   50
      TabIndex        =   10
      Top             =   2040
      Width           =   12820
      Begin VB.TextBox TxtStockTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   390
         Locked          =   -1  'True
         TabIndex        =   15
         ToolTipText     =   "stock Total (Total del material que se encuentra en el deposito)"
         Top             =   345
         Width           =   1695
      End
      Begin VB.TextBox TXTreservado 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   5445
         Locked          =   -1  'True
         TabIndex        =   14
         ToolTipText     =   "Total del Material en Fabricado en Proceso de Secado"
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox TXTCOMPRESTANTE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   8100
         Locked          =   -1  'True
         TabIndex        =   13
         ToolTipText     =   "Cantidad Vendida Pendidente de Entrega"
         Top             =   345
         Width           =   1695
      End
      Begin VB.TextBox TXTDISPONIBILIDAD 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   10905
         Locked          =   -1  'True
         TabIndex        =   12
         ToolTipText     =   "Stock Seco menos el Vendido"
         Top             =   345
         Width           =   1695
      End
      Begin VB.TextBox TxtStockSeco 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   2715
         Locked          =   -1  'True
         TabIndex        =   11
         ToolTipText     =   "Total del Material Seco"
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Total:"
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
         Left            =   120
         TabIndex        =   20
         Top             =   120
         Width           =   1425
      End
      Begin VB.Label LABEL5 
         BackStyle       =   0  'Transparent
         Caption         =   "Stock en Secado"
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
         Left            =   5115
         TabIndex        =   19
         Top             =   120
         Width           =   1425
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Vendido"
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
         Left            =   7755
         TabIndex        =   18
         Top             =   120
         Width           =   2145
      End
      Begin VB.Label label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Disponible"
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
         Left            =   10635
         TabIndex        =   17
         Top             =   120
         Width           =   1425
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Seco:"
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
         Left            =   2475
         TabIndex        =   16
         Top             =   120
         Width           =   1425
      End
   End
   Begin VB.TextBox TXTOF 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Left            =   11400
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   12920
      Begin VB.CommandButton Command1 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2775
         TabIndex        =   8
         Top             =   245
         Width           =   195
      End
      Begin VB.TextBox TXTCODIGO 
         Appearance      =   0  'Flat
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
         Left            =   1060
         TabIndex        =   7
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox TXTDESCRIPCION 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   2980
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   240
         Width           =   9735
      End
      Begin VB.Label Label3 
         BackColor       =   &H00D8BD05&
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         TabIndex        =   9
         Top             =   285
         Width           =   825
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   975
      Left            =   0
      TabIndex        =   2
      ToolTipText     =   "Cick Sobre Cantidades - Visualiza a que Ordenes de Fabricación Corresponde"
      Top             =   1080
      Width           =   12915
      _ExtentX        =   22781
      _ExtentY        =   1720
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      WordWrap        =   -1  'True
      FocusRect       =   2
      AllowUserResizing=   1
      FormatString    =   "|>CAMPO1|>CAMPO2|>CAMPO3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   3135
      Left            =   0
      TabIndex        =   21
      ToolTipText     =   "Cick Sobre Cantidades - Visualiza a que Ordenes de Fabricación Corresponde"
      Top             =   3120
      Width           =   12915
      _ExtentX        =   22781
      _ExtentY        =   5530
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      WordWrap        =   -1  'True
      FocusRect       =   2
      AllowUserResizing=   1
      FormatString    =   "|>CAMPO1|>CAMPO2|>CAMPO3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Ordenes de Fabricación con en Proceso de Secado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   45
      TabIndex        =   4
      Top             =   2910
      Width           =   4530
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Stock a Fecha de Secado para el Código Activo"
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
      Left            =   0
      TabIndex        =   3
      Top             =   860
      Width           =   4425
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
   End
   Begin VB.Menu mnuacc 
      Caption         =   "Acceso Directos"
      Begin VB.Menu mnuFikardex 
         Caption         =   "Ficha Kardex Código Activo"
      End
   End
End
Attribute VB_Name = "SITUSTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public SELECCCIONADO%
Public TXTOF1$
Private Sub Command5_Click()
    
End Sub

Private Sub Command1_Click()
  Command1.Enabled = False
  Call SELECHO("MASTER")
   COD$ = TRIM$(VALACT1$("MASTER"))
   If COD$ <> "" Then
     Me.TXTCODIGO = COD$
   End If
   Command1.Enabled = True
End Sub

Private Sub Form_Load()
   If SELECCCIONADO% > 0 Then
      Call BLOQUEAR_CONTROLES(SITUSTO, 2, 1, 1)
   End If
   
   TXTOF1$ = "" ' PARAMETRO DE LA ORDEN DE FABRICACION SELECCIONADA
   
   MSF.Rows = 3
   MSF.Clear
   DIASECADO = XVALO(CONTROL("CONFABPE", "DIASSECA"))
   If DIASECADO <= 0 Then DIASECADO = 15
   HOYX% = HOY(HO$)
   MSF.Cols = DIASECADO + 1
   MSF.ColWidth(0) = 1750
   For i& = 0 To MSF.Cols - 1
        J% = i&
        MSF.Cols = J% + 1
        MMXX1$ = METEXTO$(HOYX%)
        If J% = 0 Then
          MSF.TextMatrix(0, J%) = "HOY"
        Else
          MSF.TextMatrix(0, J%) = MMXX1$
        End If
        
        MSF.ColWidth(J%) = 750:   MSF.ColAlignment(J%) = flexAlignCenterTop
        MSF.TextMatrix(1, J%) = Left$(DSEM$(HOYX%), 3) & "-" & Left$(FECHATEX$(HOYX%), 2) & Space(10) & SAFETEXT$(HOYX%) 'GUARDO DE MANERA OCULTA LA FECHANUM
        If DiaNoLaborable%(HOYX%) > 0 Then
          Call COLOREAR_TEXTO_SOLO_UNA_AFILA(MSF, vbRed, 1, J%)
        End If
        HOYX% = DIASPOST(HOYX%, 1)
   Next i&
   '
   MSF.MergeRow(0) = True
   MSF.MergeCells = flexMergeFree
   MSF.FixedRows = 2
   
   CAMPOS$ = "CI/A0;O.Fab/A12;CÓDIGO/A20;Cantidad/F12.1;F.Cierre/D8"
   Call CARGA_ENCABEZADO(SITUSTO.MSF1, CAMPOS$, SITUSTO)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If Cancel = 0 And SELECCCIONADO% > 0 Then
      Cancel = 1: Hide
   End If
End Sub

Private Sub mnuFikardex_Click()
   FIKARDEX09.Text1 = Me.TXTCODIGO
   FIKARDEX09.Show 1
End Sub

Private Sub MSF_Click()
   REG& = MSF.MouseRow
   COL& = MSF.MouseCol
   TXTOF1$ = "" 'INICIALIZO
   TIEMPOSEC% = XVALO(CONTROL("CONFABPE", "DIASSECA"))
   If TIEMPOSEC% <= 0 Then TIEMPOSEC% = 15

   If REG& < 1 Or REG& + 1 >= MSF.Rows Then Exit Sub
   '
   CONTENIDO$ = TRIM$(MSF.TextMatrix(REG& + 1, COL&))
   FE% = Right(TRIM$(MSF.TextMatrix(1, COL&)), 6)
   FEC$ = FECHATEX(FE%)
   Call COLOREAR_GRILLASQL_TODAS_MISMO_COLOR(MSF1, VBWITH)
   If CONTENIDO$ <> "" Then
      Call TEXTOLOTES$(CONTENIDO$, ";")
      For J& = 1 To UBound(CADENATEX$)
        NORFA$ = CADENATEX$(J&)
        For i& = 1 To MSF1.Rows - 1
           FEMSF1% = FECHANUM(MSF1.TextMatrix(i&, 1)): If DIASPOST(FEMSF1%, TIEMPOSEC%) < HOY(HO$) Then FEMSF1% = DIANTE(HOY(HO$), TIEMPOSEC%)
           If TRIM$(MSF1.TextMatrix(i&, 2)) = TRIM$(NORFA$) And FE% = DIASPOST(FEMSF1%, TIEMPOSEC%) Then
              Call COLOREAR_GRILLA_SOLO_UNA_AFILA(MSF1, &HE0E0E0, i&)
           End If
        Next i&
      Next J&
   End If
   
End Sub

Private Sub MSF1_Click()
 REG& = MSF1.MouseRow
 COL& = MSF1.MouseCol
 
' If REG& > 0 Then
'    Me.TXTCodigo = MSF1.TextMatrix(REG&, 1)
'    Me.TXTCantidad = MSF1.TextMatrix(REG&, 3)
'
' End If
End Sub

Sub TXTCantidad_Click()

End Sub

Sub CargarOFabEnProcesoSimple(NORFA$)
   INFINI09.NORFAEXTERNA = NORFA$
   INFINI09.Show 1
   
End Sub

Private Sub MSF1_DblClick()
 REG& = MSF1.MouseRow
 COL& = MSF1.MouseCol
 TXTOF1$ = STRINGFILA$(SITUSTO.MSF1, REG&)
 If SELECCCIONADO% > 0 Then
    Hide
 End If
End Sub

Private Sub TXTCODIGO_Change()
CI1% = CODINT%(TXTCODIGO)
   If CI1% > 0 Then
      Me.TXTDESCRIPCION = DESCRIT0$(CI1%)
      Call MOSTRARSITUACION(CI1%)
   Else
      Me.TXTDESCRIPCION = ""
   End If
   If MSF1.Rows > 0 Then
    
    MSF1.COL = 5 'xx Columan por la q quieres ordenar
    MSF1.Sort = 1 'Aplicar orden al grid
   End If
End Sub
'
Sub MOSTRARSITUACION(CI1%)
   'TENER EN CUENTA
   'AL NO MANEJARSE MAS LOS LOTES PUEDE PASAR QUE EL DETALLE DE LA GRILLA INFERIOR
   'NO COINCIDA CON LA CANTIDAD TOTAL DISPONIBLE A FUTURO,
   'POR UN LADO SI NO ESTA SECO NO DEBERIA HABER PROBLEMAS POR QUE ESTOS ESTAN EN ESE ESTADO Y NO SE DEBERIAN
   'HABER VENDIDO
   TXTDISPONIBILIDAD.ForeColor = vbBlack
   Dim CANTIX()
   Dim CADENA_O_FAB$()
   DIASECADO = XVALO(COLOR$(CI1%))
   If DIASECADO < 1 Then
      DIASECADO = XVALO(CONTROL("CONFABPE", "DIASSECA"))
   End If
   '26/04/18 12:00 hs
   'HASTA AQUI LLEGUE
   ReDim Preserve CADENA_O_FAB$(DIASECADO)
   ReDim Preserve CANTIX(DIASECADO)
'   Call ARMALILOT(CI1%, 0, 0, 0, 1, "")
   MSF1.Rows = 1
   J& = 0
   'primero carga los saldos en la grilla inferior y en un vector guarda a que fecha va a estar seco para luego agregarlo a la grilla superior
   DISPOX = 0: RESERV = 0: CAN = 0
   HOO = HOY(HO$)
   SQLX$ = "SELECT cod_inte,COD_EXTE,DOPRILAR,CANTINGRE,FechMov  FROM MOVISTO WITH(NOLOCK)"
   SQLX$ = SQLX$ + "Where COD_INTE = " & CI1%
   SQLX$ = SQLX$ + " AND CODIMOVI = 'OF'"
   SQLX$ = SQLX$ + " AND datediff(D, FECHMOV,'" & CVDAT(HOO) & "') < " & DIASECADO

   SQLX$ = SQLX$ + " ORDER BY fechmov,DOPRILAR"

   Set RS = READSET(SQLX$)
   With RS
      Do While Not .EOF
        i& = i& + 1
        MSF1.Rows = i& + 1
        MSF1.TextMatrix(i&, 1) = XVALO(!COD_INTE)
        MSF1.TextMatrix(i&, 2) = SAFETEXT$(!DOPRILAR)
        MSF1.TextMatrix(i&, 3) = SAFETEXT$(!COD_EXTE)
        MSF1.TextMatrix(i&, 4) = TRIM$(FORMATNUM$(XVALO(!CANTINGRE), "f10.1"))
        MSF1.TextMatrix(i&, 5) = FECHATEX$(CVINT(!FechMov))
        FEDIF% = DiferenciaDiasEntreFechas(CVINT(!FechMov), HOY(HO$))
        CANTIX(Abs(FEDIF% - DIASECADO)) = CANTIX(Abs(FEDIF% - DIASECADO)) + ROUND(XVALO(!CANTINGRE), 2)
        .MoveNext
      Loop
  End With
  'carga en la grilla de arriba segun stock a fecha de secado
   For i& = 0 To DIASECADO
      MSF.TextMatrix(2, i&) = "" ' PRIMERO LIMPIO EL CASILLERO Y SI TIENE VALOR EL VECTO IMPRIMO A MODO ACUMULADO
      If XVALO(CANTIX(i&)) > 0 Then
         MSF.TextMatrix(2, i&) = CANTIX(i&) + STOCKTOT_A_SECAR
      End If
      'MSF.TextMatrix(3, i&) = Space(30) & CADENA_O_FAB$(i&)
      'MSF.RowHeight(3) = 0 ' LA OCULTO
      STOCKTOT_A_SECAR = STOCKTOT_A_SECAR + CANTIX(i&)
   Next i&
   
   If CANTIX(0) > STOCKTOT Then CANTIX(0) = STOCKTOT 'ENJUAGUE POR QUE SI HAY NEGATIVO EJ REMITO LO TERMINA SUMANDO
   Me.TxtStockTotal = FORMATNUM$(STOCKACT(CI1%), "f12.2") 'STOCK TOTAL EN DEPOSITO
   Me.TxtStockSeco = FORMATNUM$(XVALO(TxtStockTotal) - Abs(STOCKTOT_A_SECAR), "f12.2")
   Me.TXTreservado = FORMATNUM$(Abs(STOCKTOT_A_SECAR), "f12.2") 'STOCK EN SECADO
   Me.TXTCOMPRESTANTE = FORMATNUM$(CANTIDAD_PEDIDA_PENDIENTE(CI1%), "f12.2") 'VENDIDO
   
   Me.TXTDISPONIBILIDAD = FORMATNUM$(XVALO(TxtStockSeco) - XVALO(TXTCOMPRESTANTE), "f12.2") ' DISPONIBLE SECO MENOS VENDIDO
   If XVALO(TXTDISPONIBILIDAD) < 0 Then TXTDISPONIBILIDAD.ForeColor = vbRed
   
   
'   asa = CANTIDAD_PEDIDA_PENDIENTE(CI1%)
'   COMP_RESTAN = CANTIDAD_PEDIDA_PENDIENTE(CI1%) - RESERV: If COMP_RESTAN <= 0.05 Then COMP_RESTAN = 0
'   Me.TXTCOMPRESTANTE = FORMATNUM$(COMP_RESTAN, "f12.2")
'   Me.TXTDISPONIBILIDAD = FORMATNUM$((XVALO(Me.TxtStockSeco) - (XVALO(TXTCOMPRESTANTE) + XVALO(Me.TXTreservado))), "f12.2")

  Exit Sub
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   For U& = 1 To ULTREG&("!LILOTSE")
         TTYY$ = REGLEIDO$("!LILOTSE", U&)
         IDLO$ = TRIM$(Mid$(TTYY$, 1, 16))
         SALOT = SaldoLote(CI1%, IDLO$)
         IDLO_ANT$ = ""
         If SALOT > 0 Then
            SQLX$ = "SELECT * FROM MOVISTO WITH(NOLOCK)"
            SQLX$ = SQLX$ + " WHERE COD_INTE = " & CI1%
            SQLX$ = SQLX$ + " AND IDENLOTE = '" & IDLO$ & "'"
            SQLX$ = SQLX$ + " AND (CODIMOVI = 'OF' OR CODIMOVI = 'AJ' OR CODIMOVI = 'RT')"
            'SQLX$ = SQLX$ + " AND CANTINGRE > 0"
            SQLX$ = SQLX$ + " "
            SQLX$ = SQLX$ + " ORDER BY IDENLOTE,FECHMOV DESC"
            Set RS = READSET(SQLX$)
            With RS
              Do While Not .EOF
                AA = SAFETEXT$(!CODIMOVI)
                'SE AGREGO ESTO PARA QUE NO TOME LA FECHA DE LOS AJUSTES NI DE LOS REMITOS, SOLO LA DE CODIMOVI = OF (ALTA DEL LOTE) 03/06/17
                If SAFETEXT$(!CODIMOVI) = "OF" Then FE% = CVINT(!FechMov)
                HOO = HOY(HO$)
                If SAFETEXT$(!CODIMOVI) = "OF" Then FEDIF% = DiferenciaDiasEntreFechas(FE%, HOO) - DIASECADO ' SOLO CALCULA SECADO SOBRE FECHA DE OF 03/06/17
                If FEDIF% > 0 Then FEDIF% = 0
                If Abs(FEDIF%) > DIASECADO Then FEDIF% = DIASECADO
                CANTI = XVALO(!CANTINGRE) - XVALO(!Cantsalid)
'                SALOT = SALOT - CANTI
                'If SALOT - CANTI <= 0 Then CANTI = SALOT''''''''''''
                ACUCANTI = ACUCANTI + CANTI
                'carga en un vector que stocks tiene para cada fecha considerando el secado
                 CANTIX(Abs(FEDIF%)) = CANTIX(Abs(FEDIF%)) + CANTI
                'guardo en el vector CADENA_O_FAB$   para cuando hago click sobre grilla superior
                'poder saber a que o.f. pertentece las contidades secas a la fecha seleccionada al final la agrego a la grilla superior.
                If InStr(CADENA_O_FAB$(Abs(FEDIF%)), IDLO$) < 1 Then CADENA_O_FAB$(Abs(FEDIF%)) = CADENA_O_FAB$(Abs(FEDIF%)) & IDLO$ & ";"
                'If FE% <> FE_ANT% Or IDLO$ <> IDLO_ANT$ Then
                If IDLO$ <> IDLO_ANT$ Then
                  'AGREGO EN LA GRILLA INFERIOR LOS DATOS DE LAS OF.
                  J& = J& + 1
                  MSF1.Rows = J& + 1
                  MSF1.TextMatrix(J&, 1) = FECHATEX$(FE%)
                  MSF1.TextMatrix(J&, 2) = IDLO$
                  MSF1.TextMatrix(J&, 3) = CI1%
                  MSF1.TextMatrix(J&, 4) = FORMATNUM$(XVALO(MSF1.TextMatrix(J&, 4)) + CANTI, "f10.2")
                    CAN = CAN + XVALO(MSF1.TextMatrix(J&, 4))
                  MSF1.TextMatrix(J&, 5) = FE%
                  CONDI$ = "Cod_Inte = " & CI1% & " AND IdenLote = '" & IDLO$ & "'"
                  CANTIRES = XVALO(VALOBADA("RESERPED", "SUM(cantres)", CONDI$, "NOMESS"))
                  CANTCONSU = XVALO(VALOBADA("RESERPED", "sum(cantcons)", CONDI$, "NOMESS"))
                  MSF1.TextMatrix(J&, 6) = FORMATNUM$(CANTIRES - CANTCONSU, "f10.2")
                  RESERV = RESERV + XVALO(MSF1.TextMatrix(J&, 6))
                  MSF1.TextMatrix(J&, 8) = FECHATEX$(DIASPOST(FE%, Int(DIASECADO)))
                Else
                    'esto lo agregue por que mas arriba se puso un filtro p/q no lea la fecha si no es un ingreso por of
                    'pero si cuando sea una of, en los casos que primero leyo un ajuste para un mismo lote le asigna la fecha.03/06/17
                    If TRIM$(MSF1.TextMatrix(J&, 1)) = "" Then
                        TIMOX$ = SAFETEXT$(!CODIMOVI)
                        If TIMOX$ = "OF" Then
                          MSF1.TextMatrix(J&, 1) = FECHATEX$(FE%)
                        End If
                    End If
                    'si es el mismo lote suma cantiad
                    MSF1.TextMatrix(J&, 4) = FORMATNUM$(XVALO(MSF1.TextMatrix(J&, 4)) + CANTI, "f10.2")
                 End If
                 MSF1.TextMatrix(J&, 7) = FORMATNUM$(XVALO(MSF1.TextMatrix(J&, 4)) - XVALO(MSF1.TextMatrix(J&, 6)), "f10.1")
                 DISPOX = DISPOX + XVALO(MSF1.TextMatrix(J&, 7))
                 IDLO_ANT$ = IDLO$
                 FE_ANT% = FE%
                .MoveNext
              Loop
            End With
         End If
   Next U&
   MSF.Rows = 4
   'carga en la grilla de arriba segun stock a fecha de secado
   For i& = 1 To DIASECADO
      MSF.TextMatrix(2, i&) = CANTIX(i&)
      MSF.TextMatrix(3, i&) = Space(30) & CADENA_O_FAB$(i&)
      MSF.RowHeight(3) = 0 ' LA OCULTO
      STOCKTOT = STOCKTOT + CANTIX(i&)
   Next i&
   
   If CANTIX(0) > STOCKTOT Then CANTIX(0) = STOCKTOT 'ENJUAGUE POR QUE SI HAY NEGATIVO EJ REMITO LO TERMINA SUMANDO
   Me.TxtStockSeco = FORMATNUM$(CANTIX(0), "f12.2")
   Me.TxtStockTotal = FORMATNUM$(STOCKACT(CI1%), "f12.2")
   Me.TXTreservado = FORMATNUM$(RESERV, "f12.2")
   
'   asa = CANTIDAD_PEDIDA_PENDIENTE(CI1%)
   COMP_RESTAN = CANTIDAD_PEDIDA_PENDIENTE(CI1%) - RESERV: If COMP_RESTAN <= 0.05 Then COMP_RESTAN = 0
   Me.TXTCOMPRESTANTE = FORMATNUM$(COMP_RESTAN, "f12.2")
   Me.TXTDISPONIBILIDAD = FORMATNUM$((XVALO(Me.TxtStockSeco) - (XVALO(TXTCOMPRESTANTE) + XVALO(Me.TXTreservado))), "f12.2")
  
End Sub

'
Sub MOSTRARSITUACION_ANT(CI1%)
   Dim CANTIX()
   Dim CADENA_O_FAB$()
   DIASECADO = XVALO(CONTROL("CONFABPE", "DIASSECA"))
   If DIASECADO < 1 Then DIASECADO = 15
   ReDim Preserve CADENA_O_FAB$(DIASECADO)
   ReDim Preserve CANTIX(DIASECADO)
   Call ARMALILOT(CI1%, 0, 0, 0, 1, "")
   MSF1.Rows = 1
   J& = 0
   'primero carga los saldos en la grilla inferior y en un vector guarda a que fecha va a estar seco para luego agregarlo a la grilla superior
   DISPOX = 0: RESERV = 0: CAN = 0
   
   For U& = 1 To ULTREG&("!LILOTSE")
         TTYY$ = REGLEIDO$("!LILOTSE", U&)
         IDLO$ = TRIM$(Mid$(TTYY$, 1, 16))
         SALOT = SaldoLote(CI1%, IDLO$)
         IDLO_ANT$ = ""
         If SALOT > 0 Then
            SQLX$ = "SELECT * FROM MOVISTO WITH(NOLOCK)"
            SQLX$ = SQLX$ + " WHERE COD_INTE = " & CI1%
            SQLX$ = SQLX$ + " AND IDENLOTE = '" & IDLO$ & "'"
            SQLX$ = SQLX$ + " AND (CODIMOVI = 'OF' OR CODIMOVI = 'AJ' OR CODIMOVI = 'RT')"
            'SQLX$ = SQLX$ + " AND CANTINGRE > 0"
            SQLX$ = SQLX$ + " "
            SQLX$ = SQLX$ + " ORDER BY IDENLOTE,FECHMOV DESC"
            Set RS = READSET(SQLX$)
            With RS
              Do While Not .EOF
                AA = SAFETEXT$(!CODIMOVI)
                'SE AGREGO ESTO PARA QUE NO TOME LA FECHA DE LOS AJUSTES NI DE LOS REMITOS, SOLO LA DE CODIMOVI = OF (ALTA DEL LOTE) 03/06/17
                If SAFETEXT$(!CODIMOVI) = "OF" Then FE% = CVINT(!FechMov)
                HOO = HOY(HO$)
                If SAFETEXT$(!CODIMOVI) = "OF" Then FEDIF% = DiferenciaDiasEntreFechas(FE%, HOO) - DIASECADO ' SOLO CALCULA SECADO SOBRE FECHA DE OF 03/06/17
                If FEDIF% > 0 Then FEDIF% = 0
                If Abs(FEDIF%) > DIASECADO Then FEDIF% = DIASECADO
                CANTI = XVALO(!CANTINGRE) - XVALO(!Cantsalid)
'                SALOT = SALOT - CANTI
                'If SALOT - CANTI <= 0 Then CANTI = SALOT''''''''''''
                ACUCANTI = ACUCANTI + CANTI
                'carga en un vector que stocks tiene para cada fecha considerando el secado
                 CANTIX(Abs(FEDIF%)) = CANTIX(Abs(FEDIF%)) + CANTI
                'guardo en el vector CADENA_O_FAB$   para cuando hago click sobre grilla superior
                'poder saber a que o.f. pertentece las contidades secas a la fecha seleccionada al final la agrego a la grilla superior.
                If InStr(CADENA_O_FAB$(Abs(FEDIF%)), IDLO$) < 1 Then CADENA_O_FAB$(Abs(FEDIF%)) = CADENA_O_FAB$(Abs(FEDIF%)) & IDLO$ & ";"
                'If FE% <> FE_ANT% Or IDLO$ <> IDLO_ANT$ Then
                If IDLO$ <> IDLO_ANT$ Then
                  'AGREGO EN LA GRILLA INFERIOR LOS DATOS DE LAS OF.
                  J& = J& + 1
                  MSF1.Rows = J& + 1
                  MSF1.TextMatrix(J&, 1) = FECHATEX$(FE%)
                  MSF1.TextMatrix(J&, 2) = IDLO$
                  MSF1.TextMatrix(J&, 3) = CI1%
                  MSF1.TextMatrix(J&, 4) = FORMATNUM$(XVALO(MSF1.TextMatrix(J&, 4)) + CANTI, "f10.1")
                    CAN = CAN + XVALO(MSF1.TextMatrix(J&, 4))
                  MSF1.TextMatrix(J&, 5) = FE%
                  CONDI$ = "Cod_Inte = " & CI1% & " AND IdenLote = '" & IDLO$ & "'"
                  CANTIRES = XVALO(VALOBADA("RESERPED", "SUM(cantres)", CONDI$, "NOMESS"))
                  CANTCONSU = XVALO(VALOBADA("RESERPED", "sum(cantcons)", CONDI$, "NOMESS"))
                  MSF1.TextMatrix(J&, 6) = FORMATNUM$(CANTIRES - CANTCONSU, "f10.1")
                  RESERV = RESERV + XVALO(MSF1.TextMatrix(J&, 6))
                  MSF1.TextMatrix(J&, 8) = FECHATEX$(DIASPOST(FE%, Int(DIASECADO)))
                Else
                    'esto lo agregue por que mas arriba se puso un filtro p/q no lea la fecha si no es un ingreso por of
                    'pero si cuando sea una of, en los casos que primero leyo un ajuste para un mismo lote le asigna la fecha.03/06/17
                    If TRIM$(MSF1.TextMatrix(J&, 1)) = "" Then
                        TIMOX$ = SAFETEXT$(!CODIMOVI)
                        If TIMOX$ = "OF" Then
                          MSF1.TextMatrix(J&, 1) = FECHATEX$(FE%)
                        End If
                    End If
                    'si es el mismo lote suma cantiad
                    MSF1.TextMatrix(J&, 4) = FORMATNUM$(XVALO(MSF1.TextMatrix(J&, 4)) + CANTI, "f10.1")
                 End If
                 MSF1.TextMatrix(J&, 7) = FORMATNUM$(XVALO(MSF1.TextMatrix(J&, 4)) - XVALO(MSF1.TextMatrix(J&, 6)), "f10.1")
                 DISPOX = DISPOX + XVALO(MSF1.TextMatrix(J&, 7))
                 IDLO_ANT$ = IDLO$
                 FE_ANT% = FE%
                .MoveNext
              Loop
            End With
         End If
   Next U&
   MSF.Rows = 4
   'carga en la grilla de arriba segun stock a fecha de secado
   For i& = 0 To DIASECADO
      MSF.TextMatrix(2, i&) = CANTIX(i&)
      MSF.TextMatrix(3, i&) = Space(30) & CADENA_O_FAB$(i&)
      MSF.RowHeight(3) = 0 ' LA OCULTO
      STOCKTOT = STOCKTOT + CANTIX(i&)
   Next i&
   If CANTIX(0) > STOCKTOT Then CANTIX(0) = STOCKTOT 'ENJUAGUE POR QUE SI HAY NEGATIVO EJ REMITO LO TERMINA SUMANDO
   Me.TxtStockSeco = FORMATNUM$(CANTIX(0), "f12.2")
   Me.TxtStockTotal = FORMATNUM$(STOCKTOT, "f12.2")
   Me.TXTreservado = FORMATNUM$(RESERV, "f12.2")
   
'   asa = CANTIDAD_PEDIDA_PENDIENTE(CI1%)
   COMP_RESTAN = CANTIDAD_PEDIDA_PENDIENTE(CI1%) - RESERV: If COMP_RESTAN <= 0.05 Then COMP_RESTAN = 0
   Me.TXTCOMPRESTANTE = FORMATNUM$(COMP_RESTAN, "f12.2")
   Me.TXTDISPONIBILIDAD = FORMATNUM$((XVALO(Me.TxtStockSeco) - (XVALO(TXTCOMPRESTANTE) + XVALO(Me.TXTreservado))), "f12.2")
  
End Sub

Sub MOSTRARSITUACION_RESPALDO(CI1%)
   Dim CANTIX()
   Dim CADENA_O_FAB$()
   DIASECADO = XVALO(CONTROL("CONFABPE", "DIASSECA"))
   If DIASECADO < 1 Then DIASECADO = 15
   ReDim Preserve CADENA_O_FAB$(DIASECADO)
   ReDim Preserve CANTIX(DIASECADO)
   Call ARMALILOT(CI1%, 0, 0, 0, 1, "")
   MSF1.Rows = 1
   J& = 0
   'primero carga los saldos en la grilla inferior y en un vector guarda a que fecha va a estar seco para luego agregarlo a la grilla superior
   DISPOX = 0: RESERV = 0: CAN = 0
   
   For U& = 1 To ULTREG&("!LILOTSE")
         TTYY$ = REGLEIDO$("!LILOTSE", U&)
         IDLO$ = TRIM$(Mid$(TTYY$, 1, 16))
         SALOT = SaldoLote(CI1%, IDLO$)
         IDLO_ANT$ = ""
         If SALOT > 0 Then
            SQLX$ = "SELECT * FROM MOVISTO WITH(NOLOCK)"
            SQLX$ = SQLX$ + " WHERE COD_INTE = " & CI1%
            SQLX$ = SQLX$ + " AND IDENLOTE = '" & IDLO$ & "'"
            SQLX$ = SQLX$ + " AND CODIMOVI = 'OF' "
            'SQLX$ = SQLX$ + " AND CANTINGRE > 0"
            SQLX$ = SQLX$ + " "
            SQLX$ = SQLX$ + " ORDER BY IDENLOTE,FECHMOV DESC"
            Set RS = READSET(SQLX$)
            With RS
              Do While Not .EOF
'                If TRIM$(!CODIMOVI) = "AJ" Then
'                  A = 1
'                End If
                FE% = CVINT(!FechMov)
                HOO = HOY(HO$)
                FEDIF% = DiferenciaDiasEntreFechas(FE%, HOO) - DIASECADO
                If FEDIF% > 0 Then FEDIF% = 0
                If Abs(FEDIF%) > DIASECADO Then FEDIF% = DIASECADO
                CANTI = XVALO(!CANTINGRE) - XVALO(!Cantsalid)
'                SALOT = SALOT - CANTI
                If SALOT - CANTI <= 0 Then CANTI = SALOT
                ACUCANTI = ACUCANTI + CANTI
                'carga en un vector que stocks tiene para cada fecha considerando el secado
                CANTIX(Abs(FEDIF%)) = CANTIX(Abs(FEDIF%)) + CANTI
                'guardo en el vector CADENA_O_FAB$   para cuando hago click sobre grilla superior
                'poder saber a que o.f. pertentece las contidades secas a la fecha seleccionada al final la agrego a la grilla superior.
                If InStr(CADENA_O_FAB$(Abs(FEDIF%)), IDLO$) < 1 Then CADENA_O_FAB$(Abs(FEDIF%)) = CADENA_O_FAB$(Abs(FEDIF%)) & IDLO$ & ";"
                'If FE% <> FE_ANT% Or IDLO$ <> IDLO_ANT$ Then
                If IDLO$ <> IDLO_ANT$ Then
                  'AGREGO EN LA GRILLA INFERIOR LOS DATOS DE LAS OF.
                  J& = J& + 1
                  MSF1.Rows = J& + 1
                  MSF1.TextMatrix(J&, 1) = FECHATEX$(FE%)
                  MSF1.TextMatrix(J&, 2) = IDLO$
                  MSF1.TextMatrix(J&, 3) = CI1%
                  MSF1.TextMatrix(J&, 4) = FORMATNUM$(XVALO(MSF1.TextMatrix(J&, 4)) + CANTI, "f10.1")
                    CAN = CAN + XVALO(MSF1.TextMatrix(J&, 4))
                  MSF1.TextMatrix(J&, 5) = FE%
                  CONDI$ = "Cod_Inte = " & CI1% & " AND IdenLote = '" & IDLO$ & "'"
                  CANTIRES = XVALO(VALOBADA("RESERPED", "SUM(cantres)", CONDI$, "NOMESS"))
                  CANTCONSU = XVALO(VALOBADA("RESERPED", "sum(cantcons)", CONDI$, "NOMESS"))
                  MSF1.TextMatrix(J&, 6) = FORMATNUM$(CANTIRES - CANTCONSU, "f10.1")
                  RESERV = RESERV + XVALO(MSF1.TextMatrix(J&, 6))
                  MSF1.TextMatrix(J&, 8) = FECHATEX$(DIASPOST(FE%, Int(DIASECADO)))
                Else
                'si es el mismo lote suma cantiad
                  MSF1.TextMatrix(J&, 4) = FORMATNUM$(XVALO(MSF1.TextMatrix(J&, 4)) + CANTI, "f10.1")
                End If
                MSF1.TextMatrix(J&, 7) = FORMATNUM$(XVALO(MSF1.TextMatrix(J&, 4)) - XVALO(MSF1.TextMatrix(J&, 6)), "f10.1")
                DISPOX = DISPOX + XVALO(MSF1.TextMatrix(J&, 7))
                IDLO_ANT$ = IDLO$
                FE_ANT% = FE%
                .MoveNext
              Loop
            End With
         End If
   Next U&
   MSF.Rows = 4
   'carga en la grilla de arriba segun stock a fecha de secado
   For i& = 0 To DIASECADO
      MSF.TextMatrix(2, i&) = CANTIX(i&)
      MSF.TextMatrix(3, i&) = Space(30) & CADENA_O_FAB$(i&)
      MSF.RowHeight(3) = 0 ' LA OCULTO
      STOCKTOT = STOCKTOT + CANTIX(i&)
   Next i&
   
   Me.TxtStockSeco = FORMATNUM$(CANTIX(0), "f12.2")
   Me.TxtStockTotal = FORMATNUM$(STOCKTOT, "f12.2")
   Me.TXTreservado = FORMATNUM$(RESERV, "f12.2")
   
   asa = CANTIDAD_PEDIDA_PENDIENTE(CI1%)
   COMP_RESTAN = CANTIDAD_PEDIDA_PENDIENTE(CI1%) - RESERV: If COMP_RESTAN <= 0.05 Then COMP_RESTAN = 0
   Me.TXTCOMPRESTANTE = FORMATNUM$(COMP_RESTAN, "f12.2")
   Me.TXTDISPONIBILIDAD = FORMATNUM$((XVALO(Me.TxtStockSeco) - (XVALO(TXTCOMPRESTANTE) + XVALO(Me.TXTreservado))), "f12.2")
  
End Sub


