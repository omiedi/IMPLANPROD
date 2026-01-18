VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FLUCAP07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Flujo de Caja Proyectado"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   6765
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   4440
      MaskColor       =   &H00808080&
      TabIndex        =   17
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   12
      Top             =   420
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00624E28&
      ForeColor       =   &H80000008&
      Height          =   4020
      Left            =   5880
      ScaleHeight     =   3990
      ScaleWidth      =   1170
      TabIndex        =   4
      Top             =   0
      Width           =   1200
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   120
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   15
         Top             =   2280
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -120
            TabIndex        =   16
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
         Height          =   690
         Left            =   120
         Picture         =   "FLUCAP07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Genera Reporte Solicitado"
         Top             =   2640
         Width           =   650
      End
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
         Height          =   690
         Left            =   120
         Picture         =   "FLUCAP07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Tabla de Cuentas Monetarias"
         Top             =   1490
         Width           =   650
      End
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
         Height          =   690
         Left            =   120
         Picture         =   "FLUCAP07.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Acceso a Tabla de Ingresos y Gastos "
         Top             =   800
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
         Left            =   120
         Picture         =   "FLUCAP07.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   105
         Width           =   650
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2D3C0&
      Caption         =   "Opciones de Procesamiento - Incluir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3330
      Left            =   200
      TabIndex        =   0
      Top             =   120
      Width           =   5580
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Detalle de Cartera de Valores"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   6
         Left            =   600
         TabIndex        =   11
         Top             =   840
         Value           =   1  'Checked
         Width           =   4515
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Saldos Acreedores de Proveedores - Vencidos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   600
         TabIndex        =   10
         Top             =   1560
         Value           =   1  'Checked
         Width           =   4365
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Saldos Deudores de Clientes - Vencidos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   600
         TabIndex        =   9
         Top             =   1200
         Value           =   1  'Checked
         Width           =   4365
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Obligaciones Pendientes de Pagar "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   4
         Left            =   600
         TabIndex        =   7
         Top             =   2280
         Value           =   1  'Checked
         Width           =   3735
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Facturas Pendientes de Cobranza"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   3
         Left            =   600
         TabIndex        =   3
         Top             =   1920
         Value           =   1  'Checked
         Width           =   4575
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Estimacion de Gastos e Ingresos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   5
         Left            =   600
         TabIndex        =   2
         Top             =   2640
         Value           =   1  'Checked
         Width           =   4515
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Saldos Cuentas Monetarias"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   600
         TabIndex        =   1
         Top             =   480
         Value           =   1  'Checked
         Width           =   4365
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "FLUCAP07.frx":08A8
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   60
      OleObjectBlob   =   "FLUCAP07.frx":0ADC
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Menu mnuarchivo 
      Caption         =   "Archivo"
   End
   Begin VB.Menu mnuConf 
      Caption         =   "Configuracion"
   End
End
Attribute VB_Name = "FLUCAP07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub CALCULAR_IMPUESTO_RESPA(DES%, HAS%, IDCABEZAX)
    IDL = XVALO(Control("", "CTAREPER"))
    DESC$ = Control("", "DENREPER")
    
    CANTIMPUESTOS% = CantRegistros&("DATASQL", " ID_LISTA = " & IDL & " And LISTA = '" & DESC$ & "'", "NOMESS", , "CAMPO2")
    If CANTIMPUESTOS% < 1 Then Exit Sub
    '
    
    MESACTUAL% = Month(Date)
    ANOACTUAL% = Year(Date)
    If MESACTUAL% > 1 Then
       MesACalcImp% = MESACTUAL% - 1
       AnoACalcImp% = ANOACTUAL%
    Else
       MesACalcImp% = 12
       AnoACalcImp% = ANOACTUAL% - 1
    End If
    MESACTUA& = MESACTUAL%
    ANOACTUA& = ANOACTUAL%
    
    'AGREGA EL IMPORTE A PAGAR POR INGRESOS BRUTOS
    IDL2 = Control("", "CTAFVENC")
    DESC2$ = Control("", "DENFVENC")

    CODIIMPUIIBB% = XVALO(Left(Control("CASHFLOW", "CTACASIB"), 8))
    If CODIIMPUIIBB% > 0 And YAPASO% < 1 Then

       DIAVENCIM% = VALOBADA("DATASQL", "CAMPO1", "ID_LISTA = " & IDL2 & " And LISTA = '" & DESC2$ & "' AND CAST(CAMPO2 AS INTEGER) = " & CODIIMPUIIBB%, "NOMESS")
       DIAVENCI& = DIAVENCIM%
       FECHPAGO% = FECHANUM(FORMATOCON0$(DIAVENCI&, 2) & "/" & FORMATOCON0$(MESACTUA&, 2) & "/" & FORMATOCON0$(ANOACTUA&, 2))
       If FECHPAGO% > HAS% Then GoTo 50
       
       DESCRIMPU$ = TRIM$(Mid(Control("CASHFLOW", "CTACASIB"), 10, 48))
       SALDO# = IMPORTEAPAGAR_IIBB(MesACalcImp%, AnoACalcImp%)
      
       If Abs(SALDO#) > 0.005 Then
          DOCORIG$ = "IIBB x Ventas"
          DESCMOV$ = "Importe x Tasa Configurada"
          FEMI$ = FETEXCOR1$(FECHPAGO%)
          FeVencI$ = FETEXCOR1$(FECHPAGO%)
          NCLIEN = 0
          NPROV = 0
          MON% = 1
          COTIDOL# = 1
          CCBLEXLEYENDABANCO$ = CUTACBLESEGUNLEYENDA$(DESCRIPCIONFILA$)
          Call AGREGACASHFLOW(IDCABEZAX, NCLIEN, NPROV, DESCRIMPU$, SALDO#, FEMI$, FeVencI$, DOCORIG$, DESCMOV$, MON%, COTIDOL#, DESCRIMPU$, 1300, 0, 0)
        End If
    End If
    'FIN AGREGA EL IMPORTE A PAGAR POR INGRESOS BRUTOS

    
    'LEO LA TABLA DE FECHA DE VENCIMIENTOS DE IMPUESTOS
50  IDL2 = Control("", "CTAFVENC")
    DESC2$ = Control("", "DENFVENC")
    SQLX$ = "SELECT CAMPO1, CAMPO2,CAMPO3 FROM DATASQL WITH(NOLOCK) WHERE ID_LISTA = " & IDL2 & " And LISTA = '" & DESC2$ & "'"
    SQLX$ = SQLX$ + " order by campo2"

    Set rs = READSET(SQLX$)
    With rs
      Do While Not .EOF
        DIAVENCIM% = XVALO(!CAMPO1): DIAVENCI& = DIAVENCIM%
        CODIMPU% = XVALO(!CAMPO2)
        DESCRIMPU$ = SAFETEXT$(!CAMPO3)
        FECHPAGO% = FECHANUM(FORMATOCON0$(DIAVENCI&, 2) & "/" & FORMATOCON0$(MESACTUA&, 2) & "/" & FORMATOCON0$(ANOACTUA&, 2))
        If FECHPAGO% > HAS% Then GoTo 60
        
        'RECORRO LAS CUENTAS CONTABLES ASOCIADAS A CADA IMPUESTO
        SQLX$ = "SELECT CAMPO2, CAMPO3,CAMPO4, CAMPO5,CAMPO6, CAMPO7 FROM DATASQL WITH(NOLOCK) "
        SQLX$ = SQLX$ + " WHERE ID_LISTA = " & IDL & " And LISTA = '" & DESC$ & "'"
        SQLX$ = SQLX$ + " AND CAST(CAMPO2 AS INTEGER) = '" & CODIMPU% & "'"
        SQLX$ = SQLX$ + " order by campo2"
        Set RS1 = READSET(SQLX$)
        With RS1
              Do While Not .EOF
                J& = J& + 1
    '            DIAVENCIM% = XVALO(!CAMPO1)
                CODIMPU% = XVALO(!CAMPO2)
                DESCRIMPU$ = SAFETEXT$(!CAMPO3)
                CUENTACONTABLEMAS$ = SAFETEXT$(!CAMPO4)
                CUCONTMAS% = CUECOINT%(CUENTACONTABLEMAS$)
                'SI CUCONTMAS% > 0 THEN CONSULTA SQL Y GUARDAR DETALLE SOLO DEL MES ANTERIOR AL ACTUAL
                If CUCONTMAS% > 0 Then
                   GoSub GUARDARDATOSIMPUESTOS
                End If
                
                DESCRIPCIONCTACBLEMAS$ = SAFETEXT$(!CAMPO5)
                CUENTACONTABLEMENOS$ = SAFETEXT$(!CAMPO6)
                CUCONTMENOS% = CUECOINT%(CUENTACONTABLEMENOS$)
                DESCRIPCIONCTACBLEMENOS$ = SAFETEXT$(!CAMPO7)
                
                If CUCONTMENOS% > 0 Then  'CONSULTA SQL Y GUARDAR DETALLE SOLO DEL MES ANTERIOR AL ACTUAL EN NEGATIVO
                   GoSub GUARDARDATOSIMPUESTOS
                End If
                .MoveNext
              Loop
        End With
60      .MoveNext
   
      Loop
    End With

Exit Sub

GUARDARDATOSIMPUESTOS:
               If CUCONTMAS% > 0 Then
                  CUECONTABLE% = CUCONTMAS%
                  TIMOVIAMOSTRAR$ = DESCRIMPU$ & "-" & CUENTACONTABLEMAS$
               End If
               If CUCONTMENOS% > 0 Then
                  CUECONTABLE% = CUCONTMENOS%
                  TIMOVIAMOSTRAR$ = DESCRIPCIONCTACBLEMENOS$ & "-" & CUENTACONTABLEMENOS$
               End If
               SQLX$ = "SELECT  Idebecomp,ihabecomp,codicom_lar,DESCRIMOV,FECHEMISI,fechacred,NUCLIEN,NUPROVE,Mone_Emis,COTIZA_DOLAR,ValAbsComp"
               SQLX$ = SQLX$ + " FROM CAJABANC WITH(NOLOCK) WHERE CUECONTAI = " & CUECONTABLE%
               SQLX$ = SQLX$ + " and OPCIMOVIM <> 'OP'"
               SQLX$ = SQLX$ + " AND MONTH(FECHEMISI)= " & MesACalcImp%
               SQLX$ = SQLX$ + " AND YEAR(FECHEMISI)  = " & AnoACalcImp%
               Set RS2 = READSET(SQLX$)
               With RS2
               Do While Not .EOF
                  IDE# = XVALO(!IDEBECOMP)
                  IHA# = XVALO(!IHABECOMP)
                  SALDO# = XVALO(IHA# - IDE#)
                  DOCORIG$ = SAFETEXT$(!codicom_lar)
                  NCLIEN = XVALO(!NUCLIEN)
                  NPROV = XVALO(!NUPROVE)
                  'ESTO PARA QUE ACOMODE SI ES IVA DEL PROVEEDOR O SI ES IVA DE CLIENTE Y RESPETE LAS CUENTAS QUE SUMAN O RESTAN CONFIGURADAS.
                  If Abs(SALDO#) > 0.005 Then
                      If CUCONTMENOS% > 0 Then
                         SALDO# = (-1) * SALDO#
                      End If
                      If NPROV > 0 Or UCase$(Left$(DOCORIG$, 3)) = "TC-" Then
                         SALDO# = (-1) * SALDO#
                      End If
                      DESCMOV$ = SAFETEXT$(!DESCRIMOV)
                      FEMI$ = FETEXCOR1$(CVINT(!FECHEMISI))
                      FeVencI$ = FETEXCOR1$(FECHPAGO%)
                      MON% = XVALO(!Mone_Emis)
                      COTIDOL# = XVALO(!COTIZA_DOLAR)
                      Call AGREGACASHFLOW(IDCABEZAX, NCLIEN, NPROV, DESCRIMPU$, SALDO#, FEMI$, FeVencI$, DOCORIG$, DESCMOV$, MON%, COTIDOL#, DESCRIMPU$ & "-" & CUENTACONTABLEMAS$, 1300)
                  End If
                  .MoveNext
               Loop
               End With
               RS2.Close
               Set RS2 = Nothing
               '
               SQLX$ = " Select D.NUIDASIEN, D.RefePase,D.CUECOINT,D.IDebePase,D.CUECOEXT,D.IHabePase,E.TIPOASIEN, D.FECHASIEN"
               SQLX$ = SQLX$ + " FROM DetaAsien D WITH(NOLOCK)  INNER JOIN EncaBeAsien E ON D.NUIDASIEN = E.NUIDASIEN "
               SQLX$ = SQLX$ + " WHERE (E.TIPOASIEN < 1 OR E.TIPOASIEN IS NULL) "
               SQLX$ = SQLX$ + " AND D.CUECOINT = " & CUECONTABLE%
               SQLX$ = SQLX$ + " AND MONTH(D.FECHASIEN)= " & MesACalcImp%
               SQLX$ = SQLX$ + " AND YEAR(D.FECHASIEN)  = " & AnoACalcImp%
               SQLX$ = SQLX$ + " AND E.STATASIEN >= 0"
               
               Set RS3 = READSET(SQLX$)
               With RS3
               Do While Not .EOF
                  'VALIDO QUE EL TIPO DE ASIENTO SEA <1 ' NO LO AGREGUE A LA CONSULTA P Q INNER JOIN ME DUPLICA LOS REGISTROS
                  DOCORIG$ = SAFETEXT$(!nuidasien)
                  CALIFICA& = AsientoTieneCuentaMadreProveedor&(DOCORIG$, CUECONTABLE%) 'DEBE EXISTIR COMO  CUENTA MADRE DE PROVEEDOR
                  If CALIFICA& < 1 Then GoTo 85 'NEXT 'AsientoTieneCuentaMadreProveedor& DEVUELVE EL NRO DE PROVEEDOR
                  IDE# = XVALO(!idebepase)
                  IHA# = XVALO(!ihabepase)
                  SALDO# = XVALO(IDE# - IHA#)
                  CUEXT$ = SAFETEXT$(!CUECOEXT)
                  NCLIEN = 0
                  NPROV = CALIFICA&
                  'ESTO PARA QUE ACOMODE SI ES IVA DEL PROVEEDOR O SI ES IVA DE CLIENTE Y RESPETE LAS CUENTAS QUE SUMAN O RESTAN CONFIGURADAS.
                  If Abs(SALDO#) > 0.005 Then

                      DESCMOV$ = SAFETEXT$(!RefePase)
                      FEMI$ = FETEXCOR1$(CVINT(!FECHASIEN))
                      FeVencI$ = FETEXCOR1$(FECHPAGO%)
                      MON% = 1
                      COTIDOL# = 1
                      Call AGREGACASHFLOW(IDCABEZAX, NCLIEN, NPROV, DESCRIMPU$, SALDO#, FEMI$, FeVencI$, DOCORIG$, DESCMOV$, MON%, COTIDOL#, DESCRIMPU$ & "-" & CUEXT$, 1300)
                  End If
85                .MoveNext
               Loop
               End With
               RS3.Close
               Set RS3 = Nothing

               CUCONTMAS% = 0: CUCONTMENOS% = 0
               
Return
    
End Sub

Sub CALCULAR_IMPUESTO(DES%, HAS%, IDCABEZAX)
    IDL = XVALO(Control("", "CTAREPER"))
    DESC$ = Control("", "DENREPER")
    
    CANTIMPUESTOS% = CantRegistros&("DATASQL", " ID_LISTA = " & IDL & " And LISTA = '" & DESC$ & "'", "NOMESS", , "CAMPO2")
    If CANTIMPUESTOS% < 1 Then Exit Sub
    '
    MESACTUAL% = Month(Date)
    ANOACTUAL% = Year(Date)
    If MESACTUAL% > 1 Then
       MesACalcImp% = MESACTUAL% - 1
       AnoACalcImp% = ANOACTUAL%
    Else
       MesACalcImp% = 12
       AnoACalcImp% = ANOACTUAL% - 1
    End If
    MESACTUA& = MESACTUAL%
    ANOACTUA& = ANOACTUAL%
    
    'AGREGA EL IMPORTE A PAGAR POR INGRESOS BRUTOS
    IDL2 = Control("", "CTAFVENC")
    DESC2$ = Control("", "DENFVENC")
    CODIIMPUIIBB% = XVALO(Left(Control("CASHFLOW", "CTACASIB"), 8))
    If CODIIMPUIIBB% > 0 And YAPASO% < 1 Then
       DIAVENCIM% = XVALO(VALOBADA("DATASQL", "CAMPO1", "ID_LISTA = " & IDL2 & " And LISTA = '" & DESC2$ & "' AND CAST(CAMPO2 AS INTEGER) = " & CODIIMPUIIBB%, "NOMESS"))
       DIAVENCI& = DIAVENCIM%
       FECHPAGO% = FECHANUM(FORMATOCON0$(DIAVENCI&, 2) & "/" & FORMATOCON0$(MESACTUA&, 2) & "/" & FORMATOCON0$(ANOACTUA&, 2))
'       FEMISION% = FECHPAGO%)
       F1$ = HOY(HO$)
       'CAMBIO SOLICITADO SI HOY ES ANTERIOR A LA FECHA DE VENCIMIENTO CALCULA EL MES ANTERIOR COMO ANTES
       'SI ES HOY ES POSTERIOR AL DIA DE VENCIMIENTO CALCULA EL MES CORRIENTE Y LA FECHA DE VENCIMIENTO VA A SER PARA EL MES SIGUIENTE
       If FECHPAGO% < HOY(HO$) Then
          MesACalcImp% = MESACTUAL%: MESACTUA& = MESACTUAL%
          AnoACalcImp% = ANOACTUAL%: ANOACTUA& = ANOACTUAL%
          MESACTUA& = MESACTUA& + 1
          If MESACTUA& > 12 Then 'PARA CALCULAR LA FechPago% DEL VENCIMIENTO AL MES SIGUIENTE.
             MESACTUA& = 1
             ANOACTUA& = ANOACTUA& + 1
          End If
          FECHPAGO% = FECHANUM(FORMATOCON0$(DIAVENCI&, 2) & "/" & FORMATOCON0$(MESACTUA&, 2) & "/" & FORMATOCON0$(ANOACTUA&, 2))
       End If
       'FIN INTERVENCION DE CAMBIO SOLICITADO
       
       If FECHPAGO% > HAS% Then GoTo 50

       DESCRIMPU$ = TRIM$(Mid(Control("CASHFLOW", "CTACASIB"), 10, 48))
       CUCONTABANCO$ = CUTACBLESEGUNLEYENDA$(DESCRIMPU$)
       SALDO# = IMPORTEAPAGAR_IIBB(MesACalcImp%, AnoACalcImp%)

       If Abs(SALDO#) > 0.005 Then
          DOCORIG$ = "IIBB x Ventas"
          DESCMOV$ = "Importe x Tasa Configurada"
          FEMI$ = FETEXCOR1$(FECHPAGO%)
          FeVencI$ = FETEXCOR1$(FECHPAGO%)
          NCLIEN = 0
          NPROV = 0
          MON% = 1
          COTIDOL# = 1
          Call AGREGACASHFLOW(IDCABEZAX, NCLIEN, NPROV, DESCRIMPU$, SALDO#, FEMI$, FeVencI$, DOCORIG$, DESCMOV$, MON%, COTIDOL#, DESCRIMPU$, 1300, 0, "", CUCONTABANCO$)
        End If
    End If
    'FIN AGREGA EL IMPORTE A PAGAR POR INGRESOS BRUTOS

    
    'LEO LA TABLA DE FECHA DE VENCIMIENTOS DE IMPUESTOS
50  IDL2 = Control("", "CTAFVENC")
    DESC2$ = Control("", "DENFVENC")
    SQLX$ = "SELECT CAMPO1, CAMPO2,CAMPO3 FROM DATASQL WITH(NOLOCK) WHERE ID_LISTA = " & IDL2 & " And LISTA = '" & DESC2$ & "'"
    SQLX$ = SQLX$ + " order by campo2"
    Set rs = READSET(SQLX$)
    With rs
      Do While Not .EOF
        DIAVENCIM% = XVALO(!CAMPO1): DIAVENCI& = DIAVENCIM%
        CODIMPU% = XVALO(!CAMPO2)
        DESCRIMPU$ = SAFETEXT$(!CAMPO3)
        MESACTUA& = MESACTUAL%: ANOACTUA& = ANOACTUAL%
    
        FECHPAGO% = FECHANUM(FORMATOCON0$(DIAVENCI&, 2) & "/" & FORMATOCON0$(MESACTUA&, 2) & "/" & FORMATOCON0$(ANOACTUA&, 2))
       'CAMBIO SOLICITADO SI HOY ES ANTERIOR A LA FECHA DE VENCIMIENTO CALCULA EL MES ANTERIOR COMO ANTES
       'SI ES HOY ES POSTERIOR AL DIA DE VENCIMIENTO CALCULA EL MES CORRIENTE Y LA FECHA DE VENCIMIENTO VA A SER PARA EL MES SIGUIENTE
       If FECHPAGO% < HOY(HO$) Then
          MesACalcImp% = MESACTUAL%: MESACTUA& = MESACTUAL%
          AnoACalcImp% = ANOACTUAL%: ANOACTUA& = ANOACTUAL%
          MESACTUA& = MESACTUA& + 1
          If MESACTUA& > 12 Then 'PARA CALCULAR LA FechPago% DEL VENCIMIENTO AL MES SIGUIENTE.
             MESACTUA& = 1
             ANOACTUA& = ANOACTUA& + 1
          End If
          FECHPAGO% = FECHANUM(FORMATOCON0$(DIAVENCI&, 2) & "/" & FORMATOCON0$(MESACTUA&, 2) & "/" & FORMATOCON0$(ANOACTUA&, 2))
       End If
       'FIN INTERVENCION DE CAMBIO SOLICITADO

        
        If FECHPAGO% > HAS% Then GoTo 60
        
        'RECORRO LAS CUENTAS CONTABLES ASOCIADAS A CADA IMPUESTO
        SQLX$ = "SELECT CAMPO2, CAMPO3,CAMPO4, CAMPO5,CAMPO6, CAMPO7 FROM DATASQL WITH(NOLOCK) "
        SQLX$ = SQLX$ + " WHERE ID_LISTA = " & IDL & " And LISTA = '" & DESC$ & "'"
        SQLX$ = SQLX$ + " AND CAST(CAMPO2 AS INTEGER) = '" & CODIMPU% & "'"
        SQLX$ = SQLX$ + " order by campo2"
        Set RS1 = READSET(SQLX$)
        With RS1
              Do While Not .EOF
                J& = J& + 1
    '            DIAVENCIM% = XVALO(!CAMPO1)
                CODIMPU% = XVALO(!CAMPO2)
                DESCRIMPU$ = SAFETEXT$(!CAMPO3)
                CUCONTABANCO$ = CUTACBLESEGUNLEYENDA$(DESCRIMPU$)

                CUENTACONTABLEMAS$ = SAFETEXT$(!CAMPO4)
                CUCONTMAS% = CUECOINT%(CUENTACONTABLEMAS$)
                'SI CUCONTMAS% > 0 THEN CONSULTA SQL Y GUARDAR DETALLE SOLO DEL MES ANTERIOR AL ACTUAL
                If CUCONTMAS% > 0 Then
                   GoSub GUARDARDATOSIMPUESTOS
                End If
                
                DESCRIPCIONCTACBLEMAS$ = SAFETEXT$(!CAMPO5)
                CUENTACONTABLEMENOS$ = SAFETEXT$(!CAMPO6)
                CUCONTMENOS% = CUECOINT%(CUENTACONTABLEMENOS$)
                DESCRIPCIONCTACBLEMENOS$ = SAFETEXT$(!CAMPO7)
                
                If CUCONTMENOS% > 0 Then  'CONSULTA SQL Y GUARDAR DETALLE SOLO DEL MES ANTERIOR AL ACTUAL EN NEGATIVO
                   GoSub GUARDARDATOSIMPUESTOS
                End If
                .MoveNext
              Loop
        End With
60      .MoveNext
   
      Loop
    End With

Exit Sub

GUARDARDATOSIMPUESTOS:
               If CUCONTMAS% > 0 Then
                  CUECONTABLE% = CUCONTMAS%
                  TIMOVIAMOSTRAR$ = DESCRIMPU$ & "-" & CUENTACONTABLEMAS$
               End If
               If CUCONTMENOS% > 0 Then
                  CUECONTABLE% = CUCONTMENOS%
                  TIMOVIAMOSTRAR$ = DESCRIPCIONCTACBLEMENOS$ & "-" & CUENTACONTABLEMENOS$
               End If
               FILA$ = 1
               SQLX$ = "SELECT  Idebecomp,ihabecomp,codicom_lar,DESCRIMOV,FECHEMISI,fechacred,NUCLIEN,NUPROVE,Mone_Emis,COTIZA_DOLAR,ValAbsComp"
               SQLX$ = SQLX$ + " FROM CAJABANC WITH(NOLOCK) WHERE CUECONTAI = " & CUECONTABLE%
               SQLX$ = SQLX$ + " and OPCIMOVIM <> 'OP'"
               SQLX$ = SQLX$ + " AND MONTH(FECHEMISI)= " & MesACalcImp%
               SQLX$ = SQLX$ + " AND YEAR(FECHEMISI)  = " & AnoACalcImp%
               Set RS2 = READSET(SQLX$)
               With RS2
               FILA$ = 2
               Do While Not .EOF
                  IDE# = XVALO(!IDEBECOMP)
                  IHA# = XVALO(!IHABECOMP)
                  SALDO# = XVALO(IHA# - IDE#)
                  DOCORIG$ = SAFETEXT$(!codicom_lar)
                  NCLIEN = XVALO(!NUCLIEN)
                  NPROV = XVALO(!NUPROVE)
                  'ESTO PARA QUE ACOMODE SI ES IVA DEL PROVEEDOR O SI ES IVA DE  CLIENTE Y RESPETE LAS CUENTAS QUE SUMAN O RESTAN CONFIGURADAS.
                  If Abs(SALDO#) > 0.005 Then
'                      If CUCONTMENOS% > 0 Then
'                         SALDO# = (-1) * SALDO#
'                      End If
'                      If NPROV > 0 Or UCase$(Left$(DOCORIG$, 3)) = "TC-" Then
'                         SALDO# = (-1) * SALDO#
'                      End If
                      DESCMOV$ = SAFETEXT$(!DESCRIMOV)
                      FEMI$ = FETEXCOR1$(CVINT(!FECHEMISI))
                      FeVencI$ = FETEXCOR1$(FECHPAGO%)
                      MON% = XVALO(!Mone_Emis)
                      COTIDOL# = XVALO(!COTIZA_DOLAR)
                      Call AGREGACASHFLOW(IDCABEZAX, NCLIEN, NPROV, DESCRIMPU$, SALDO#, FEMI$, FeVencI$, DOCORIG$, DESCMOV$, MON%, COTIDOL#, DESCRIMPU$ & "-" & CUENTACONTABLEMAS$, 1300, 0, "", CUCONTABANCO$)
                  End If
                  .MoveNext
               Loop
               End With
               RS2.Close
               Set RS2 = Nothing
               '
               FILA$ = 3
               SQLX$ = " Select D.NUIDASIEN, D.RefePase,D.CUECOINT,D.IDebePase,D.CUECOEXT,D.IHabePase,E.TIPOASIEN, D.FECHASIEN"
               SQLX$ = SQLX$ + " FROM DetaAsien D WITH(NOLOCK)  INNER JOIN EncaBeAsien E ON D.NUIDASIEN = E.NUIDASIEN "
               SQLX$ = SQLX$ + " WHERE (E.TIPOASIEN < 1 OR E.TIPOASIEN IS NULL) "
               SQLX$ = SQLX$ + " AND D.CUECOINT = " & CUECONTABLE%
               SQLX$ = SQLX$ + " AND MONTH(D.FECHASIEN)= " & MesACalcImp%
               SQLX$ = SQLX$ + " AND YEAR(D.FECHASIEN)  = " & AnoACalcImp%
               SQLX$ = SQLX$ + " AND E.STATASIEN >= 0 AND (MARCONSI = '' or MARCONSI is null)"
               
               Set RS3 = READSET(SQLX$)
               FILA$ = 4
               With RS3
               Do While Not .EOF
                  'VALIDO QUE EL TIPO DE ASIENTO SEA <1 ' NO LO AGREGUE A LA CONSULTA P Q INNER JOIN ME DUPLICA LOS REGISTROS
                  DOCORIG$ = SAFETEXT$(!nuidasien)
                  CALIFICA& = AsientoTieneCuentaMadreProveedor&(DOCORIG$, CUECONTABLE%) 'DEBE EXISTIR COMO  CUENTA MADRE DE PROVEEDOR
                  If CALIFICA& < 1 Then GoTo 85 'NEXT 'AsientoTieneCuentaMadreProveedor& DEVUELVE EL NRO DE PROVEEDOR
                  IDE# = XVALO(!idebepase)
                  IHA# = XVALO(!ihabepase)
                  SALDO# = XVALO(IDE# - IHA#)
                  CUEXT$ = SAFETEXT$(!CUECOEXT)
                  NCLIEN = 0
                  NPROV = CALIFICA&
                  'ESTO PARA QUE ACOMODE SI ES IVA DEL PROVEEDOR O SI ES IVA DE CLIENTE Y RESPETE LAS CUENTAS QUE SUMAN O RESTAN CONFIGURADAS.
                  If Abs(SALDO#) > 0.005 Then

                      DESCMOV$ = SAFETEXT$(!RefePase)
                      FEMI$ = FETEXCOR1$(CVINT(!FECHASIEN))
                      FeVencI$ = FETEXCOR1$(FECHPAGO%)
                      MON% = 1
                      COTIDOL# = 1
                      Call AGREGACASHFLOW(IDCABEZAX, NCLIEN, NPROV, DESCRIMPU$, SALDO#, FEMI$, FeVencI$, DOCORIG$, DESCMOV$, MON%, COTIDOL#, DESCRIMPU$ & "-" & CUEXT$, 1300)
                  End If
85                .MoveNext
               Loop
               End With
               RS3.Close
               FILA$ = 5
               Set RS3 = Nothing

               CUCONTMAS% = 0: CUCONTMENOS% = 0
            On Error GoTo 0
               
Return
    
End Sub



Sub ELIMINACASHFLOW()
    If DERACCE%("DELCASHF", "Eliminación de Cash-Flow Procesado") < 2 Then Exit Sub
    
10  campos$ = "Nro. /F12;Referencia/A64;Fecha y Hora/A20,FDES/A0;FHAS/A0;INTVLO/A0"
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, campos$, FRMZELECHO)
    FRMZELECHO.Caption = "Cashflow Emitidos"

    Dim obj As New RECORXET
    Set RS1 = obj.RS_CashFlowEmitidos
    Call Carga_MSF_Recordset(RS1, campos$, FRMZELECHO.MSF2, 1)
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    FRMZELECHO.Width = 12000
    FRMZELECHO.Show 1
    If XVALO(RESP_ZELE_VECT(1)) > 0 Then
      Screen.MousePointer = 11
      ID_CABEZA& = XVALO(RESP_ZELE_VECT(1))
      FLEXCONN.BeginTrans
      Call BORRAREGISTROS("ENCAFLOW", "ENCAFLOW_ID = " & ID_CABEZA&, REGAF&, "NOMESS")
      Call BORRAREGISTROS("BACKUPFLOW", "ID_CABEZERA = " & ID_CABEZA&, REGAF&, "NOMESS")
      Call T_Espera(1)
      FLEXCONN.CommitTrans
      Screen.MousePointer = 1
      GoTo 10
    End If
    
End Sub

Sub GENERAR_CABEZA_CASHFLOW(TEXTO1$, TEXTO2$, RET_IDENTITY, DESDE%, HASTA%, INTERVALO$)
        'DA DE ALTA UN NUEVO IDCABEZA CON LA REFERENCIA DE LA MODALIDAD Y EL RANGO DE FECHAS Y RETORNA EL IDENTITY CREADO
        SQLX$ = "INSERT INTO ENCAFLOW "
        SQLX$ = SQLX$ + " (REFERENCIA1,REFERENCIA2,FECHAYHO,FEDES,FEHAS,INTERVALO)"
        SQLX$ = SQLX$ + " VALUES('" & TEXTO1$ & "'"
        SQLX$ = SQLX$ + " ,'" & TEXTO2$ & "'"
        SQLX$ = SQLX$ + " , GETDATE()  "
        SQLX$ = SQLX$ + " ,'" & CVDAT(DESDE%) & "'"
        SQLX$ = SQLX$ + " ,'" & CVDAT(HASTA%) & "'"
        SQLX$ = SQLX$ + " ,'" & INTERVALO & "'"
        
        
        SQLX$ = SQLX$ + "); SELECT SCOPE_IDENTITY();" 'devuelve el ultimo valor de identity
        FLEXCONN.Execute (SQLX$)
        '
        'RETORNA EL IDENTITY CREADO
        SQLX$ = " SELECT SCOPE_IDENTITY() AS NUMCOL "
        Set rs = FLEXCONN.Execute(FILTSQL$(SQLX$))
        RET_IDENTITY = rs!NUMCOL
        '
End Sub

Function COEFI_IMPOCONCOMISION#(NOMBTARJETA$)
   'LEE DE LA TABLA DATASQL SEGUN EL NOMBRE DEL MEDIO DE PAGO Y LO COMPARA CON EL DESCRIMOV QUE ALLI ESTA GUARDADO
   'EN LOS 18 PRIMEROS CARACTERES EL MISMO
   'COMO EL MOVIMIENTO ES TC. LEE COMPARA A PARTIR DE LA 4 POSICION
   Static IDL, DESC$
   'EJ. SI EL PORCENTAJE DE COMISION ES 10 GUARDADO
   If IDL < 1 Then
     IDL = XVALO(Control("", "CTAACCOM"))
     DESC$ = Control("", "DENACCOM")
   End If
   Condi$ = " ID_LISTA = " & IDL & " And LISTA = '" & DESC$ & "' AND RTRIM(LTRIM(SUBSTRING(CAMPO2,4,LEN(CAMPO2)-3)))=  '" & NOMBTARJETA$ & "'"
   PorcentajeX = XVALO(VALOBADA("DATASQL", "CAMPO4", Condi$, "NOMESS"))
   If PorcentajeX >= 99.99 Or PorcentajeX <= 0 Then
      COEFI_IMPOCONCOMISION# = 1
   Else
      COEFI_IMPOCONCOMISION# = (100 - PorcentajeX) / 100
   End If

End Function

Function IMPORTEAPAGAR_IIBB(MES, ANO)
    'DEVUELVE EL IMPORTE A PAGAR SEGUN LAS VENTAS DE UN MES Y AÑO CONSIDERANDO LA TASA CONFIGURADA PARA EL CASHFLOW
    Condi$ = "TIPOMOVIM = 'VENTA' AND   Month(FECHEMISI) = " & MES
    Condi$ = Condi$ + " AND YEAR(FECHEMISI)  = " & ANO
    IDEB# = XVALO(VALOBADA("CAJABANC", "Sum(IDEBECOMP)", Condi$, "NOMESS"))
    IHAB# = XVALO(VALOBADA("CAJABANC", "Sum(IHABECOMP)", Condi$, "NOMESS"))
    VENTAX = IHAB# - IDEB#
    TASAX = XVALO(Control("CASHFLOW", "TASCASIB"))
    If TASAX > 0 Then
       IMPORTEAPAGAR_IIBB = (VENTAX * TASAX) / 100
    End If
End Function

Sub INGYGASTOSFIJOS(DES%, HAS%, ID_CAB)
  IDL = XVALO(Control("", "CTAINGGAS"))
  DESC$ = TRIM$(Control("", "DENINGGAS"))

  J& = 0
  SQLX$ = "SELECT CAMPOFE1, CAMPO2,CAMPO3,CAMPO4,CAMPO5 FROM DATASQL WITH(NOLOCK) WHERE ID_LISTA = " & IDL & " And LISTA = '" & DESC$ & "'"
  SQLX$ = SQLX$ + " AND CAMPOFE1 >= '" & FETEXCOR1$(DES%) & "'"
  SQLX$ = SQLX$ + " AND CAMPOFE1 <= '" & FETEXCOR1$(HAS%) & "'"
  SQLX$ = SQLX$ + " ORDER BY CAMPOFE1"
  Set rs = READSET(SQLX$)
  With rs
    Do While Not .EOF
      J& = J& + 1
      FECHA$ = CVDAT(CVINT(!CAMPOFE1))
      Codigo$ = SAFETEXT$(!CAMPO2)
      REFERENCIA$ = SAFETEXT$(!CAMPO3)
      INGRESO# = XVALO(!CAMPO4)
      GASTO# = XVALO(!CAMPO5)
      IMPORTX# = 0
      If INGRESO# > 0 Then
         TIPOMOVIX$ = "Ingreso"
         TIPOAMOSTRAR$ = "Ingreso Fijo"
         IMPORTX# = INGRESO#
         Orden& = 900
      ElseIf GASTO# > 0 Then
         TIPOMOVIX$ = "Gasto"
         IMPORTX# = GASTO#
         TIPOAMOSTRAR$ = "Gasto"
         Orden& = 1400
      End If
      DESCRIPCIONFILA$ = REFERENCIA$
      CUCONTABANCO$ = CUTACBLESEGUNLEYENDA$(DESCRIPCIONFILA$)
      Call AGREGACASHFLOW(ID_CAB, 0, 0, REFERENCIA$, IMPORTX#, FECHA$, FECHA$, REFERENCIA$, REFERENCIA$, 1, 1, TIPOAMOSTRAR$, Orden&, 0, " ", CUCONTABANCO$)
      .MoveNext
    Loop
  End With
End Sub

Sub TOTALIZARCASHFLOW(MSF As MSFlexGrid)
   Dim SALDONETOIN#()
   Dim SALDONETOEGR#()
   ReDim SALDONETOIN#(MSF.Cols - 2)
   ReDim SALDONETOEGR#(MSF.Cols - 2)
   
   For i& = 1 To MSF.Rows - 1
     If MSF.TextMatrix(i&, 0) = "DETALLE DE INGRESOS" Then FILA_IN_INGRESOS = i&
     If MSF.TextMatrix(i&, 0) = "TOTAL INGRESOS" Then FILA_FIN_INGRESOS = i&
     
     If MSF.TextMatrix(i&, 0) = "DETALLE DE EGRESOS" Then FILA_IN_EGRESOS = i&
     If MSF.TextMatrix(i&, 0) = "TOTAL EGRESOS" Then FILA_FIN_EGRESOS = i&
   Next i&
   
   For c& = 1 To MSF.Cols - 2 '-2 POR QUE HAY UNA ULTIMA COLUMNA NO VISIBLE
       IMPO# = 0
       For i& = FILA_IN_INGRESOS To FILA_FIN_INGRESOS - 1
         IMPO# = IMPO# + XVALO(MSF.TextMatrix(i&, c&))
       Next i&

       MSF.TextMatrix(i&, c&) = FORMATNUM$(IMPO#, "F14.2")
       SALDONETOIN#(c&) = XVALO(IMPO#)
   Next c&

   For c& = 1 To MSF.Cols - 2 '-2 POR QUE HAY UNA ULTIMA COLUMNA NO VISIBLE
       IMPO# = 0
       For i& = FILA_IN_EGRESOS To FILA_FIN_EGRESOS - 1
         IMPO# = IMPO# + XVALO(MSF.TextMatrix(i&, c&))
       Next i&
       MSF.TextMatrix(i&, c&) = FORMATNUM$(IMPO#, "F14.2")
       SALDONETOEGR#(c&) = XVALO(IMPO#)
   Next c&
   '
   'AGREGO 2 FILAS MAS
   MSF.Rows = MSF.Rows + 2
   MSF.TextMatrix(MSF.Rows - 2, 0) = "SALDO NETO"
   MSF.TextMatrix(MSF.Rows - 1, 0) = "SALDO ACUMULADO"
   Call COLOREAR_GRILLA_SOLO_UNA_AFILA(MSF, &HC0C0C0, MSF.Rows - 2)
   Call COLOREAR_GRILLA_SOLO_UNA_AFILA(MSF, &H808080, MSF.Rows - 1)
   'SALDO NETO - 'SALDO ACUMULADO
   REG& = MSF.Rows - 2
   For c& = 1 To MSF.Cols - 2
      MSF.TextMatrix(REG&, c&) = FORMATNUM$(XVALO(SALDONETOIN#(c&)) - XVALO(SALDONETOEGR#(c&)), "F14.2")
      If c& = 1 Then
         MSF.TextMatrix(REG& + 1, c&) = MSF.TextMatrix(REG&, c&)
      Else
         MSF.TextMatrix(REG& + 1, c&) = XVALO(MSF.TextMatrix(REG& + 1, c& - 1)) + XVALO(MSF.TextMatrix(REG&, c&))
      End If
   Next c&
   
   

End Sub

Private Sub BOTREC_Click()
    '
    BOTREC.Enabled = False
    '
    For i& = 1 To Check1.Count
       Check1(i& - 1).Enabled = False
    Next i&
    '
    COMESA$ = FETEXCOR1$(COMESACT%)
    
    If InStr(EMPREAC$, "DOSIVAC") > 0 Then
        Call CASHFLOW
        BOTREC.Enabled = True
        Exit Sub
    End If
    
    If Check1(0).Value = 1 Then
       'Lee saldos de cuentas monetarias
15     JJ& = 0: CXI% = 0
       For i& = 1 To ULTREG("TACUMONE")
         XX$ = REGLEIDO$("TACUMONE", i&)
         CC$ = TRIM(Left$(XX$, 12))
         COCUE% = CUEINT%(CC$)
         DESCUE$ = TRIM$(Mid(XX$, 13, 48))
         '
         HOII% = HOY(HO$)
         SALDO# = SARRASCUE#(COCUE%, HOII%): DEBE# = 0: HABER# = 0
         If SALDO# < 0 Then
              HABER# = (-1) * SALDO#
            Else
              DEBE# = SALDO#
         End If
         '
         JJ& = JJ& + 1
         YYY$ = REGBLAN("FLUCAPRO")
         Call REPLA(YYY$, MKI(HOY(HO$)), 1, 2) 'FECHA HOY
         Call REPLA(YYY$, "A.SALDO", 3, 12)
         Call REPLA(YYY$, CC$, 15, 12)
         Call REPLA(YYY$, DESCUE$, 27, 30)
         Call REPLA(YYY$, MKD(DEBE#), 57, 8)
         Call REPLA(YYY$, MKD(HABER#), 65, 8)
         Call GRAREG("FLUCAPRO", YYY$, JJ&)
       Next i&
       Call CIERRARCH("TACUMONE")
       'Fin lectura saldos de cuentas monetarias
    End If
    '
    If Check1(1).Value = 1 Then
       ' Saldos Deudores de Clientes - Vencidos
20      SQLX$ = "SELECT NUCLIEN,sum(SALDDEBE) AS SUDEB, SUM(SALDACRE) AS SUHAB FROM SADEUPEN WITH(NOLOCK)"
        SQLX$ = SQLX$ + " WHERE FEVENCIM < '" & COMESA$ & "' "
        SQLX$ = SQLX$ + " GROUP BY NUCLIEN "
        '
        Dim SADEUTMP As ADODB.Recordset
        Set SADEUTMP = READSET(SQLX$)
        With SADEUTMP
           Do While .EOF = False
              YYY$ = REGBLAN("FLUCAPRO")
              FECVEN% = HOY(HOS$)
              DEUTO# = XVALO(!SUDEB) - XVALO(!SUHAB)
              If Abs(DEUTO#) > 0.005 Then
                 Call REPLA(YYY$, MKI(FECVEN%), 1, 2)
                    NC1 = XVALO(!NUCLIEN)
                 Call REPLA(YYY$, "B.DEU-VENCI", 3, 12)
                 Call REPLA(YYY$, CODICLI$(NC1), 15, 12)
                 Call REPLA(YYY$, rasocli$(NC1), 27, 30)
                 If DEUTO# > 0 Then
                      Call REPLA(YYY$, MKD(DEUTO#), 57, 8)
                   Else
                      Call REPLA(YYY$, MKD(Abs(DEUTO#)), 65, 8)
                 End If
                 Call REPLA(YYY$, MKD(0), 53, 8)
                 '
                 JJ& = JJ& + 1
                 Call GRAREG("FLUCAPRO", YYY$, JJ&)
              End If
           .MoveNext
           Loop
        End With
        SADEUTMP.Close
        Set SADEUTMP = Nothing
    End If

    If Check1(2).Value = 1 Then
       ' Saldos Acreedores de Proveedores - Vencidos
21     SQLX$ = "SELECT NUPROVE,sum(SALDDEBE) AS SUDEB, SUM(SALDACRE) AS SUHAB FROM SACREPEN WITH(NOLOCK)"
        SQLX$ = SQLX$ + " WHERE FEVENCIM < '" & COMESA$ & "' "
        SQLX$ = SQLX$ + " GROUP BY NUPROVE "
        '
        Set SADEUTMP = READSET(SQLX$)
        With SADEUTMP
           Do While .EOF = False
              YYY$ = REGBLAN("FLUCAPRO")
              FECVEN% = HOY(HOS$)
              DEUTO# = XVALO(!SUHAB) - XVALO(!SUDEB)
              If Abs(DEUTO#) > 0.005 Then
                 Call REPLA(YYY$, MKI(FECVEN%), 1, 2)
                    NC1 = (-1) * XVALO(!NUPROVE)
                 Call REPLA(YYY$, "C.PRO-VENCI", 3, 12)
                 Call REPLA(YYY$, CODICLI$(NC1), 15, 12)
                 Call REPLA(YYY$, rasocli$(NC1), 27, 30)
                 If DEUTO# < 0 Then
                      Call REPLA(YYY$, MKD(Abs(DEUTO#)), 57, 8)
                   Else
                      Call REPLA(YYY$, MKD(DEUTO#), 65, 8)
                 End If
                 Call REPLA(YYY$, MKD(0), 53, 8)
                 '
                 JJ& = JJ& + 1
                 Call GRAREG("FLUCAPRO", YYY$, JJ&)
              End If
           .MoveNext
           Loop
        End With
        SADEUTMP.Close
        Set SADEUTMP = Nothing
    End If

    If Check1(3).Value = 1 Then
        SQLX$ = "SELECT NUCLIEN,FEVENCIM,CODICOM_COR, SALDDEBE AS SUDEB, SALDACRE AS SUHAB FROM SADEUPEN WITH(NOLOCK)"
        SQLX$ = SQLX$ + " WHERE FEVENCIM >= '" & COMESA$ & "' "
        '
        Set SADEUTMP = READSET(SQLX$)
        With SADEUTMP
           Do While .EOF = False
              YYY$ = REGBLAN("FLUCAPRO")
              FECVEN% = MAXIVAL(CVINT(!FEVENCIM), HOY(HOS$))
              DEUTO# = XVALO(!SUDEB) - XVALO(!SUHAB)
              If Abs(DEUTO#) > 0.005 Then
                 Call REPLA(YYY$, MKI(FECVEN%), 1, 2)
                    NC1 = XVALO(!NUCLIEN)
                 Call REPLA(YYY$, SAFETEXT(!CodiCom_Cor), 3, 12)
                 Call REPLA(YYY$, CODICLI$(NC1), 15, 12)
                 Call REPLA(YYY$, rasocli$(NC1), 27, 30)
                 If DEUTO# > 0 Then
                      Call REPLA(YYY$, MKD(DEUTO#), 57, 8)
                   Else
                      Call REPLA(YYY$, MKD(Abs(DEUTO#)), 65, 8)
                 End If
                 Call REPLA(YYY$, MKD(0), 53, 8)
                 '
                 JJ& = JJ& + 1
                 Call GRAREG("FLUCAPRO", YYY$, JJ&)
              End If
           .MoveNext
           Loop
        End With
        SADEUTMP.Close
        Set SADEUTMP = Nothing
    End If
    '
    If Check1(4).Value = 1 Then
        SQLX$ = "SELECT NUPROVE,FEVENCIM,CODICOM_COR, SALDDEBE AS SUDEB, SALDACRE AS SUHAB FROM SACREPEN WITH(NOLOCK)"
        SQLX$ = SQLX$ + " WHERE FEVENCIM >= '" & COMESA$ & "' "
        '
        Set SADEUTMP = READSET(SQLX$)
        With SADEUTMP
           Do While .EOF = False
              YYY$ = REGBLAN("FLUCAPRO")
              FECVEN% = MAXIVAL(CVINT(!FEVENCIM), HOY(HO$))
              DEUTO# = XVALO(!SUHAB) - XVALO(!SUDEB)
              If Abs(DEUTO#) > 0.005 Then
                 Call REPLA(YYY$, MKI(FECVEN%), 1, 2)
                    NC1 = (-1) * XVALO(!NUPROVE)
                 Call REPLA(YYY$, SAFETEXT(!CodiCom_Cor), 3, 12)
                 Call REPLA(YYY$, CODICLI$(NC1), 15, 12)
                 Call REPLA(YYY$, rasocli$(NC1), 27, 30)
                 If DEUTO# < 0 Then
                      Call REPLA(YYY$, MKD(Abs(DEUTO#)), 57, 8)
                   Else
                      Call REPLA(YYY$, MKD(DEUTO#), 65, 8)
                 End If
                 Call REPLA(YYY$, MKD(0), 53, 8)
                 '
                 JJ& = JJ& + 1
                 Call GRAREG("FLUCAPRO", YYY$, JJ&)
              End If
           .MoveNext
           Loop
        End With
        SADEUTMP.Close
        Set SADEUTMP = Nothing
    End If
    '
    If Check1(5).Value = 1 Then
       SQLX$ = "SELECT * FROM PLAGASPEN WITH(NOLOCK)"
       Dim PLAGTMPE As ADODB.Recordset
       Set PLAGTMPE = READSET(SQLX$)
       With PLAGTMPE
          Do While Not .EOF
             '
             INGGAS$ = "INGRESO"
             If XVALO(!GASTOS) > 0.005 Then
                INGGAS$ = "GASTO"
             End If
             '
             YYY$ = REGBLAN("FLUCAPRO")
             FE% = CVINT(!FECHA)
             Call REPLA(YYY$, MKI$(FE%), 1, 2)
             Call REPLA(YYY$, INGGAS$, 3, 12)
             Call REPLA(YYY$, (SAFETEXT$(!REFERENCIA)), 27, 30)
             Call REPLA(YYY$, MKD$(XVALO(!INGRESOS)), 57, 8)
             Call REPLA(YYY$, MKD$(XVALO(!GASTOS)), 65, 8)
             JJ& = JJ& + 1
             Call GRAREG("FLUCAPRO", YYY$, JJ&)
             '
          .MoveNext   '
          Loop
       End With
       PLAGTMPE.Close
       Set PLAGTMPE = Nothing
    End If
    '
    If Check1(6).Value = 1 Then
        'Grabacion de los cheques
50      SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK)"
        SQLX$ = SQLX$ + "where TipoMovim = 'CHRECI' "
        SQLX$ = SQLX$ + "AND Status = 1 "
        SQLX$ = SQLX$ + "ORDER BY FechAcred ASC;"
        '
        Dim CABATEMP As ADODB.Recordset
        Set CABATEMP = READSET(SQLX$)
        With CABATEMP
           Do While Not (.EOF)
             FECACRE% = MAXIVAL(CVINT(!fechacred), DIPOST(HOY(HO$)))
             CAR$ = "CARTERA"
             NURECIB$ = SAFETEXT(!CodiCom_Cor)
             NUCHEQUE$ = Mid(SAFETEXT(!DESCRIMOV), 1, 30)
             '
             YYY$ = REGBLAN("FLUCAPRO")
             '
             Call REPLA(YYY$, MKI(FECACRE%), 1, 2)
             Call REPLA(YYY$, CAR$, 3, 12)
             Call REPLA(YYY$, NURECIB$, 15, 12)
             Call REPLA(YYY$, NUCHEQUE$, 27, 30)
             Call REPLA(YYY$, MKD(XVALO(!IDEBECOMP)), 57, 8)
             '
             JJ& = JJ& + 1
             Call GRAREG("FLUCAPRO", YYY$, JJ&)
           .MoveNext
           Loop
        End With
        '
    End If
    '
    '
90  Call SETULTREG("FLUCAPRO", JJ&)
    Call CIERRARCH("FLUCAPRO")
    Call FILESORT("FLUCAPRO", "", 1, 2, "ASC")
    SARRA# = 0: FANTE% = 0
    FIN& = JJ&
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("FLUCAPRO", U&)
      If CVI(Left$(XX$, 2)) = FANTE% Then
        If U& > 1 Then
          YY$ = REGLEIDO$("FLUCAPRO", U& - 1)
          Call REPLA(YY$, MKD$(0), 73, 8)
          Call GRAREG("FLUCAPRO", YY$, U& - 1)
          XX$ = REGLEIDO$("FLUCAPRO", U&)
        End If
      End If
      FANTE% = CVI(Left$(XX$, 2))
      SARRA# = SARRA# + CVD(Mid$(XX$, 57, 8)) - CVD(Mid$(XX$, 65, 8))
      Call REPLA(XX$, MKD$(SARRA#), 73, 8)
      Call GRAREG("FLUCAPRO", XX$, U&)
    Next U&
    '
95  Call CIERRARCH("FLUCAPRO")
    Call FINAL("*LIFLUCA")
    Screen.MousePointer = 1
        
99  BOTREC.Enabled = True
    For i& = 1 To Check1.Count
       Check1(i& - 1).Enabled = True
    Next i&
    '
End Sub

Sub CASHFLOW()

    Dim FLUJOCAJA(365, 12) As Double
    Call ABRECONEXION

    Call CREACAMPO("", "FLUCAJA", "FECHA", 8, 0)
    Call CREACAMPO("", "FLUCAJA", "BANCO", 7, 0)
    Call CREACAMPO("", "FLUCAJA", "EFECTIVO", 7, 0)
    Call CREACAMPO("", "FLUCAJA", "FAPECOB", 7, 0)
    Call CREACAMPO("", "FLUCAJA", "CHACRED", 7, 0)
    Call CREACAMPO("", "FLUCAJA", "CHEMITI", 7, 0)
    Call CREACAMPO("", "FLUCAJA", "FAPEPAG", 7, 0)
    Call CREACAMPO("", "FLUCAJA", "SUELDOS", 7, 0)
    Call CREACAMPO("", "FLUCAJA", "CARSOCI", 7, 0)
    Call CREACAMPO("", "FLUCAJA", "INTERBAN", 7, 0)
    Call CREACAMPO("", "FLUCAJA", "GASESTRU", 7, 0)
    Call CREACAMPO("", "FLUCAJA", "EGREVAR", 7, 0)
    Call CREACAMPO("", "FLUCAJA", "COMEXTE", 7, 0)

    For II& = 0 To 365
        For J = 0 To 12
            FLUJOCAJA(II&, J) = 0
        Next J
    Next II&
    
'    Call SELECHO("SELFECHA")
'    FECHA$ = TRIM$(VALACT1$("SELFECHA"))
'    If FECHA$ = "" Then GoTo 99
'    FECHACOM% = FECHANUM(FECHA$)
'    COMESA$ = FETEXCOR1$(FECHACOM%)
    Call DESHASFECHA(DES%, HAS%, PERIO$)
    If PERIO$ = "" Then GoTo 99
    FECHACOM% = DES%: COMFLOW$ = FETEXCOR1$(DES%)
    FECHAFIN% = HAS%: FINFLOW$ = FETEXCOR1$(HAS%)
    
    Screen.MousePointer = 11
    ' BANCOS
    Condi$ = "TIPO=29 AND FECHA>='" & COMFLOW$ & "'" & " AND FECHA<='" & FINFLOW$ & "'"
    FIN& = CantRegistros("PLAGASPEN", Condi$)
    JJ& = 1

    SQLX$ = "SELECT * FROM PLAGASPEN WITH(NOLOCK) WHERE " + Condi$
    Dim PLABANCOS As ADODB.Recordset
    Set PLABANCOS = READSET(SQLX$)
    With PLABANCOS
       Do While Not .EOF
          Call TRACE(24, JJ&, FIN&)
          FECVEN% = MAXIVAL(CVINT(!FECHA), FECHACOM%)
          SALDO# = XVALO(!INGRESOS)
          If Abs(SALDO#) > 0.005 Then
            If FECVEN% - FECHACOM% > 365 Then
               III% = 365
            Else
               III% = FECVEN% - FECHACOM%
            End If
            FLUJOCAJA(III%, 1) = FLUJOCAJA(III%, 1) + SALDO#
          End If
          JJ& = JJ& + 1
       .MoveNext
       Loop
    End With
    PLABANCOS.Close
    Set PLABANCOS = Nothing
    
    ' EFECTIVO
    FF1% = FECHACOM% - 1
    FIN& = ULTREG("TACUMONE")
    JJ& = 1
    For II& = 1 To FIN&
      Call TRACE(24, JJ&, FIN&)
      XX$ = REGLEIDO$("TACUMONE", II&)
      CC$ = TRIM(Left$(XX$, 12))
      COCUE% = CUEINT%(CC$)
      '
      For JJ1& = 0 To 364
            FF1% = DIPOST(FF1%)
            If FF1% <= FECHAFIN% Then
                Condi$ = "CueContaI = " & COCUE% & " AND FECHCONTA = '" & FETEXCOR1$(FF1%) & "'"
                IDEBE# = XVALO(VALOBADA("CAJABANC", "SUM(IDEBECOMP)", Condi$))
                IHABE# = XVALO(VALOBADA("CAJABANC", "SUM(IHABECOMP)", Condi$))
                SALDO# = IDEBE# - IHABE#
                If Abs(SALDO#) > 0 Then
                    III% = JJ1&
                    FLUJOCAJA(III%, 2) = FLUJOCAJA(III%, 2) + SALDO#
                End If
            End If
      Next JJ1&
      '
      JJ& = JJ& + 1
    Next II&
    
    'FACTURAS PENDIENTES DE COBRANZA
    Condi$ = "FEVENCIM<='" + FINFLOW$ + "'"
    FIN& = CantRegistros("SADEUPEN", Condi$)
    JJ& = 1
    
    SQLX$ = "SELECT NUCLIEN,FEVENCIM,CODICOM_COR, SALDDEBE AS SUDEB, SALDACRE AS SUHAB FROM SADEUPEN WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE " + Condi$
    Dim SADEUTMP As ADODB.Recordset
    Set SADEUTMP = READSET(SQLX$)
    With SADEUTMP
       Do While .EOF = False
          Call TRACE(24, JJ&, FIN&)
          FECVEN% = MAXIVAL(CVINT(!FEVENCIM), FECHACOM%)
          DEUTO# = XVALO(!SUDEB) - XVALO(!SUHAB)
          If Abs(DEUTO#) > 0.005 Then
            If FECVEN% - FECHACOM% > 365 Then
               III% = 365
            Else
               III% = FECVEN% - FECHACOM%
            End If
            FLUJOCAJA(III%, 3) = FLUJOCAJA(III%, 3) + DEUTO#
          End If
          JJ& = JJ& + 1
       .MoveNext
       Loop
    End With
    SADEUTMP.Close
    Set SADEUTMP = Nothing
    
    ' CHEQUES A ACREDITAR
    Condi$ = "TipoMovim = 'CHRECI' AND Status = 1 AND FechAcred>='" & COMFLOW$ & "' AND FECHACRED<='" & FINFLOW$ & "'"
    FIN& = CantRegistros("CAJABANC", Condi$)
    JJ& = 1
    '
    SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE " + Condi$
    SQLX$ = SQLX$ + " ORDER BY FechAcred ASC;"
    Dim CHEQACRETEMP As ADODB.Recordset
    Set CHEQACRETEMP = READSET(SQLX$)
    With CHEQACRETEMP
       Do While Not (.EOF)
         Call TRACE(24, JJ&, FIN&)
         FECACRE% = MAXIVAL(CVINT(!fechacred), FECHACOM%)
         SALDO# = XVALO(!IDEBECOMP)
         If Abs(SALDO#) > 0.005 Then
            If FECACRE% - FECHACOM% > 365 Then
               III% = 365
            Else
               III% = FECACRE% - FECHACOM%
            End If
            FLUJOCAJA(III%, 4) = FLUJOCAJA(III%, 4) + SALDO#
         End If
         JJ& = JJ& + 1
         .MoveNext
       Loop
    End With
    CHEQACRETEMP.Close
    Set CHEQACRETEMP = Nothing

    ' CHEQUES A DEBITAR
    If ULTREG&("CUEBANC") > 0 Then
        Condi$ = "TipoMovim = 'CHENTR' "
        Condi$ = Condi$ + "AND FechAcred>='" & COMFLOW$ & "' AND FechAcred<='" & FINFLOW$ & "' AND ("
        For IJ& = 1 To ULTREG&("CUEBANC")
            XX$ = REGLEIDO$("CUEBANC", IJ&)
            CCON% = CUEINT%(Mid$(XX$, 1, 12))
            Condi$ = Condi$ + "CueContaI = " & CCON% & " OR "
        Next IJ&
        Condi$ = Mid(Condi$, 1, Len(Condi$) - 3) + ")"
        FIN& = CantRegistros("CAJABANC", Condi$)
        JJ& = 0
        '
        SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK)"
        SQLX$ = SQLX$ + "WHERE " + Condi$
        SQLX$ = SQLX$ + " ORDER BY FechAcred ASC;"
        Dim CHEQDEBITEMP As ADODB.Recordset
        Set CHEQDEBITEMP = READSET(SQLX$)
        With CHEQDEBITEMP
           Do While Not (.EOF)
             Call TRACE(24, JJ&, FIN&)
             FECDEBI% = CVINT(!fechacred)
             SALDO# = XVALO(!IHABECOMP)
             If Abs(SALDO#) > 0.005 Then
                If FECDEBI% - FECHACOM% > 365 Then
                   III% = 365
                Else
                   III% = FECDEBI% - FECHACOM%
                End If
                FLUJOCAJA(III%, 5) = FLUJOCAJA(III%, 5) + SALDO#
             End If
             JJ& = JJ& + 1
             .MoveNext
           Loop
        End With
        CHEQDEBITEMP.Close
        Set CHEQDEBITEMP = Nothing
    End If

    ' FACTURAS PENDIENTES DE PAGO
    Condi$ = "FEVENCIM<='" + FINFLOW$ + "'"
    FIN& = CantRegistros("SACREPEN", Condi$)
    JJ& = 0
    '
    SQLX$ = "SELECT NUPROVE,FEVENCIM,CODICOM_COR, SALDDEBE AS SUDEB, SALDACRE AS SUHAB FROM SACREPEN WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE " + Condi$
    Set ACRETMP = READSET(SQLX$)
    With ACRETMP
       Do While .EOF = False
          Call TRACE(24, JJ&, FIN&)
          FECVEN% = MAXIVAL(CVINT(!FEVENCIM), FECHACOM%)
          SALDO# = XVALO(!SUHAB) - XVALO(!SUDEB)
          If Abs(SALDO#) > 0.005 Then
            If FECVEN% - FECHACOM% > 365 Then
               III% = 365
            Else
               III% = FECVEN% - FECHACOM%
            End If
            FLUJOCAJA(III%, 6) = FLUJOCAJA(III%, 6) + SALDO#
          End If
          JJ& = JJ& + 1
       .MoveNext
       Loop
    End With
    ACRETMP.Close
    Set ACRETMP = Nothing
    
    ' PLAN DE SUELDOS
    Condi$ = "TIPO=1 AND FECHA>='" & COMFLOW$ & "'" & " AND FECHA<='" & FINFLOW$ & "'"
    FIN& = CantRegistros("PLAGASPEN", Condi$)
    JJ& = 0
    '
    SQLX$ = "SELECT * FROM PLAGASPEN WITH(NOLOCK) WHERE " + Condi$
    Dim PLAGASUEL As ADODB.Recordset
    Set PLAGASUEL = READSET(SQLX$)
    With PLAGASUEL
       Do While Not .EOF
          Call TRACE(24, JJ&, FIN&)
          FECVEN% = MAXIVAL(CVINT(!FECHA), FECHACOM%)
          SALDO# = XVALO(!GASTOS)
          If Abs(SALDO#) > 0.005 Then
            If FECVEN% - FECHACOM% > 365 Then
               III% = 365
            Else
               III% = FECVEN% - FECHACOM%
            End If
            FLUJOCAJA(III%, 7) = FLUJOCAJA(III%, 7) + SALDO#
          End If
          JJ& = JJ& + 1
       .MoveNext
       Loop
    End With
    PLAGASUEL.Close
    Set PLAGASUEL = Nothing

    ' PLAN DE CARGAS SOCIALES
    Condi$ = "TIPO=2 AND FECHA>='" & COMFLOW$ & "'" & " AND FECHA<='" & FINFLOW$ & "'"
    FIN& = CantRegistros("PLAGASPEN", Condi$)
    JJ& = 0
    '
    SQLX$ = "SELECT * FROM PLAGASPEN WITH(NOLOCK) WHERE " + Condi$
    Dim PLAGACARSO As ADODB.Recordset
    Set PLAGACARSO = READSET(SQLX$)
    With PLAGACARSO
       Do While Not .EOF
          Call TRACE(24, JJ&, FIN&)
          FECVEN% = MAXIVAL(CVINT(!FECHA), FECHACOM%)
          SALDO# = XVALO(!GASTOS)
          If Abs(SALDO#) > 0.005 Then
            If FECVEN% - FECHACOM% > 365 Then
               III% = 365
            Else
               III% = FECVEN% - FECHACOM%
            End If
            FLUJOCAJA(III%, 8) = FLUJOCAJA(III%, 8) + SALDO#
          End If
          JJ& = JJ& + 1
       .MoveNext
       Loop
    End With
    PLAGACARSO.Close
    Set PLAGACARSO = Nothing

    ' PLAN DE INTERBANKING
    Condi$ = "TIPO=3 AND FECHA>='" & COMFLOW$ & "'" & " AND FECHA<='" & FINFLOW$ & "'"
    FIN& = CantRegistros("PLAGASPEN", Condi$)
    JJ& = 0
    '
    SQLX$ = "SELECT * FROM PLAGASPEN WITH(NOLOCK) WHERE " + Condi$
    Dim PLAGAINTER As ADODB.Recordset
    Set PLAGAINTER = READSET(SQLX$)
    With PLAGAINTER
       Do While Not .EOF
          Call TRACE(24, JJ&, FIN&)
          FECVEN% = MAXIVAL(CVINT(!FECHA), FECHACOM%)
          SALDO# = XVALO(!GASTOS)
          If Abs(SALDO#) > 0.005 Then
            If FECVEN% - FECHACOM% > 365 Then
               III% = 365
            Else
               III% = FECVEN% - FECHACOM%
            End If
            FLUJOCAJA(III%, 9) = FLUJOCAJA(III%, 9) + SALDO#
          End If
          JJ& = JJ& + 1
       .MoveNext
       Loop
    End With
    PLAGAINTER.Close
    Set PLAGAINTER = Nothing

    ' PLAN DE GASTOS DE ESTRUCTURA
    Condi$ = "TIPO=4 AND FECHA>='" & COMFLOW$ & "'" & " AND FECHA<='" & FINFLOW$ & "'"
    FIN& = CantRegistros("PLAGASPEN", Condi$)
    JJ& = 0
    '
    SQLX$ = "SELECT * FROM PLAGASPEN WITH(NOLOCK) WHERE " + Condi$
    Dim PLAGAESTRU As ADODB.Recordset
    Set PLAGAESTRU = READSET(SQLX$)
    With PLAGAESTRU
       Do While Not .EOF
          Call TRACE(24, JJ&, FIN&)
          FECVEN% = MAXIVAL(CVINT(!FECHA), FECHACOM%)
          SALDO# = XVALO(!GASTOS)
          If Abs(SALDO#) > 0.005 Then
            If FECVEN% - FECHACOM% > 365 Then
               III% = 365
            Else
               III% = FECVEN% - FECHACOM%
            End If
            FLUJOCAJA(III%, 10) = FLUJOCAJA(III%, 10) + SALDO#
          End If
          JJ& = JJ& + 1
       .MoveNext
       Loop
    End With
    PLAGAESTRU.Close
    Set PLAGAESTRU = Nothing
    
    ' PLAN DE GASTOS VARIOS
    Condi$ = "TIPO=0 AND FECHA>='" & COMFLOW$ & "'" & " AND FECHA<='" & FINFLOW$ & "'"
    FIN& = CantRegistros("PLAGASPEN", Condi$)
    JJ& = 0
    '
    SQLX$ = "SELECT * FROM PLAGASPEN WITH(NOLOCK) WHERE " + Condi$
    Dim PLAEGRESOS As ADODB.Recordset
    Set PLAEGRESOS = READSET(SQLX$)
    With PLAEGRESOS
       Do While Not .EOF
          Call TRACE(24, JJ&, FIN&)
          FECVEN% = MAXIVAL(CVINT(!FECHA), FECHACOM%)
          SALDO# = XVALO(!GASTOS)
          If Abs(SALDO#) > 0.005 Then
            If FECVEN% - FECHACOM% > 365 Then
               III% = 365
            Else
               III% = FECVEN% - FECHACOM%
            End If
            FLUJOCAJA(III%, 11) = FLUJOCAJA(III%, 11) + SALDO#
          End If
          JJ& = JJ& + 1
       .MoveNext
       Loop
    End With
    PLAEGRESOS.Close
    Set PLAEGRESOS = Nothing
    
    ' PLAN DE GASTOS
    Condi$ = "Stat_Ocom < 8 "
    Condi$ = Condi$ + " AND Stat_Ocom <> -1 "
    Condi$ = Condi$ + " AND Cant_Rec < Cant_Ocom - 0.001 "
    Condi$ = Condi$ + " AND FEntre_Sol>='" + COMFLOW$ + "' AND FEntre_Sol<='" + FINFLOW$ + "'"
    FIN& = CantRegistros("OCOMDETA", Condi$)
    JJ& = 0
    '
    SQLX$ = "Select * from OCOMDETA WITH(NOLOCK)"
    SQLX$ = SQLX$ + " Where " + Condi$
    SQLX$ = SQLX$ + " ORDER BY FEntre_Sol "
    Dim OCOMTEMP As ADODB.Recordset
    Set OCOMTEMP = READSET(SQLX$)
    With OCOMTEMP
       Do While Not .EOF
          Call TRACE(24, JJ&, FIN&)
          NPRO% = XVALO(!Npro_Ocom): NROOCOM& = XVALO(!nume_ocom)
          CONDIVA% = XVALO(VALOBADA("PROVED12", "Piva_Cli", "NUME_CLI = " & NPRO%))
          If CONDIVA% = 5 And STATOCOM%(NROOCOM&) < 3 Then
            FECVEN% = MAXIVAL(CVINT(!Fentre_Sol), FECHACOM%)
            MONEMI% = XVALO(!Mone_Emis)
            SALDO# = XVALO(!cant_ocom) * XVALO(!prun_neto) * TICAMONE(MONEMI%)
            If Abs(SALDO#) > 0.005 Then
              If FECVEN% - FECHACOM% > 365 Then
                 III% = 365
              Else
                 III% = FECVEN% - FECHACOM%
              End If
              FLUJOCAJA(III%, 12) = FLUJOCAJA(III%, 12) + SALDO#
            End If
          End If
          JJ& = JJ& + 1
       .MoveNext
       Loop
    End With
    OCOMTEMP.Close
    Set OCOMTEMP = Nothing
    
    'COMPLETO TABLA SQL CON LOS DATOS DE LA MATRIZ
    Condi$ = "FECHA > '" & FETEXCOR1$(FECHANUM("01/01/80")) & "' OR FECHA IS NULL"
    Call BORRAREGISTROS("FLUCAJA", Condi$, REGAF&)
    Dim FLUCAJATMP As ADODB.Recordset
    Set FLUCAJATMP = New ADODB.Recordset
    SQLX$ = "SELECT * FROM FLUCAJA WHERE " & Condi$
    FLUCAJATMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText

    FF% = FECHACOM% - 1
    FECHAFIN% = MINIVAL(FECHAFIN%, FECHACOM% + 365)
    With FLUCAJATMP
        For II& = 0 To 365
            FF1% = DIPOST(FF%): NN% = FF1% - FECHACOM%
            If FF1% <= FECHAFIN% Then
                .AddNew
                !FECHA = CVDAT(FF1%)
                !BANCO = !BANCO + FLUJOCAJA(NN%, 1)
                !EFECTIVO = !EFECTIVO + FLUJOCAJA(NN%, 2)
                !FAPECOB = !FAPECOB + FLUJOCAJA(NN%, 3)
                !CHACRED = !CHACRED + FLUJOCAJA(NN%, 4)
                !CHEMITI = !CHEMITI + FLUJOCAJA(NN%, 5)
                !FAPEPAG = !FAPEPAG + FLUJOCAJA(NN%, 6)
                !SUELDOS = !SUELDOS + FLUJOCAJA(NN%, 7)
                !CARSOCI = !CARSOCI + FLUJOCAJA(NN%, 8)
                !INTERBAN = !INTERBAN + FLUJOCAJA(NN%, 9)
                !GASESTRU = !GASESTRU + FLUJOCAJA(NN%, 10)
                !EGREVAR = !EGREVAR + FLUJOCAJA(NN%, 11)
                !COMEXTE = !COMEXTE + FLUJOCAJA(NN%, 12)
                .Update
            End If
            FF% = FF1%
        Next II&
    End With
    FLUCAJATMP.Close
    Set FLUCAJATMP = Nothing
    
    Screen.MousePointer = 1
    Call COMUNI("Proceso Finalizado")
    
    Call CONECRUN("?CASHFLOW", "Flujo de Caja Proyectado")
99  Exit Sub
End Sub


'
Private Sub Command1_Click()
   Call CIERRARCH("*.*")
   Label1 = ""
   Hide
   
End Sub

Private Sub Command3_Click()

    Command3.Enabled = False
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) = "" Then
        GoTo 99
    End If
    '
    MMXX0$ = TRIM$(MESASEL$)
    MESAN = MMXX0$
    '
    MMXX1$ = MESANO$(MMXX0$, DES%, HAS%)
99  Command3.Enabled = True

End Sub


Sub Command2_Click()
   Call CREATABLAS_CASHFLOW
    
   OPX% = COMALTER%("Opciones de Cash-flow\\Cancelar\Ver Cash-Flow Procesados\Emitir Nuevo Cash-flow\Configuración\Eliminar Cash-Flow Procesado")
   If OPX% <= 1 Then Exit Sub
   If OPX% = 2 Then Call VERCASHEMITIDOS: Exit Sub
   If OPX% = 4 Then FSCASHFLOW.Show 1: Exit Sub
   If OPX% = 5 Then Call ELIMINACASHFLOW: Exit Sub
   'NUEVO CASHFLOW
   If DERACCE%("GENECASH", "Generación de Cash-Flow ") < 2 Then Exit Sub
   'VALIDACIONES
   
   CODIIMPUIIBB% = XVALO(Left(Control("CASHFLOW", "CTACASIB"), 8))
   If CODIIMPUIIBB% < 1 Then
      Call COMUNI("Falta Configurar Cual es la Cuenta de Ingresos Brutos para Poder Calcular El Importe a Pagar")
      GoTo 99
   End If
   
   TASAX = XVALO(Control("CASHFLOW", "TASCASIB"))
   If TASAX <= 0 Then
      Call COMUNI("Falta Configurar Cual es la Tasa Para el Cálculo de Ingresos Brutos a Pagar")
      GoTo 99
   End If
   
   ' CHEQUES A ACREDITAR
   FEX = HOY(HO$)
   HAS% = FEX
   HASS$ = FECHATEX$(FEX)
   DESS$ = "01" + Mid$(HASS$, 3)
   DES% = FECHANUM(DESS$)
   Call REPLA(VDEF$, MKI$(DES%), 1, 2) ' INFORMACION PARA EL OBJETO
   Call REPLA(VDEF$, MKI$(HAS%), 3, 2)
   Call REPLA(VDEF$, "SEM", 5, 4)

    '
30  OBX$ = OBJPLA$("DEHACASH", VDEF$) ' ABRE EL OBJETO
    '
    If OBX$ = "" Then Exit Sub
    DES% = CVI(Mid$(OBX$, 1, 2))
    HAS% = CVI(Mid$(OBX$, 3, 2))
    MODX$ = TRIM$(Mid$(OBX$, 5, 4))
    '

    '
    If DES% > HAS% Then
       Call COMUNI("Atención La Fecha Desde No Puede Ser Mayor a la Fecha Hasta")
       GoTo 99
    End If
    '
    MODX$ = UCase$(MODX$)
    If MODX$ <> "DIA" And MODX$ <> "SEM" And MODX$ <> "MES" Then
       Call COMUNI("Atención Los Intérvalo Aceptados son: DIA, SEM, MES")
       GoTo 99
    End If
    '
    'MUESTRA LA PANTALLA PARA EL INGRESO DE LOS SALDOS BANCARIOS
    IDL = XVALO(Control("", "CTALCUBA"))
    Call ABM_TABLA(XVALO(IDL))
    Screen.MousePointer = 11
    

    Rango$ = "Emitido Día: " & FECHATEX$(HOY(HO$))
    Rango$ = Rango$ & " Rango: " & FECHATEX$(DES%) & " al " & FECHATEX$(HAS%)
    Rango$ = Rango$ & " Modo: " & MODX$
    '
    'DE ACUERDO A QUE INTERVALO SE SELECCIONO CARGA EL ENCABEZADO.
    Select Case MODX$
    Case "DIA"
        Call CARGA_ENCABEZADO_DIAS(PRESFLOW.MSF1, PRESFLOW, DES%, HAS%, "F12.2", 1, RETCANTIDAS%, RETNUMPRIMERDIA%, RETNUMULTIMDIA%, NOMBREFILAS$)
    Case "SEM"
        Call CARGA_ENCABEZADO_SEMANA(PRESFLOW.MSF1, PRESFLOW, DES%, HAS%, "F21.2", 1, RETCANTISEMANAS%, RETNUMPRIMERASEMANA%, RETNUMULTIMSEMANA%, NOMBREFILAS$)
    Case "MES"
        Call CARGA_ENCABEZADO_MESES(PRESFLOW.MSF1, PRESFLOW, DES%, HAS%, "F20.2", 1, RETCANTIMESES%, RETNUMPRIMERAMES%, RETNUMULTIMES%, NOMBREFILAS$)
    End Select
    PRESFLOW.Caption = Rango$
    'TRANSACCIÓN GRABACION
    On Error GoTo ERROR_GUARDAR
    FLEXCONN.BeginTrans
    
    Call GENERAR_CABEZA_CASHFLOW(Rango$, " ", RET_IDCABEZA, DES%, HAS%, MODX$)

    DESDEX$ = FETEXCOR1$(DES%)
    DESDEMENOS30DIAS% = DIASANTES(DES%, 30) 'SE USA PARA LOS CHEQUES PROPIOS A DEBITAR RETROCEDE EL DESDE 30 DIAS
    HASTAX$ = FETEXCOR1$(HAS%)
    '
    Call CALCULAR_IMPUESTO(DES%, HAS%, RET_IDCABEZA)
    Call INGYGASTOSFIJOS(DES%, HAS%, RET_IDCABEZA)
    Condi$ = "TipoMovim = 'CHRECI' AND Status = 1 AND FechAcred>='" & DESDEX$ & "' AND FECHACRED<='" & HASTAX$ & "'"

    FIN& = CantRegistros("CAJABANC", Condi$)
    JJ& = 1
    '
    'CARTERA DE VALORES
    CHRECIX% = 0
    SQLX$ = "SELECT Idebecomp,codicom_lar,DESCRIMOV,FECHEMISI,fechacred,NUCLIEN,Mone_Emis,COTIZA_DOLAR FROM CAJABANC WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE " + Condi$
    SQLX$ = SQLX$ + " AND OPCIMOVIM <> 'TC.' AND OPCIMOVIM <> 'TC' AND OPCIMOVIM <> 'ME' AND (MARCONSI = '' or MARCONSI is null)"
    SQLX$ = SQLX$ + " ORDER BY FechAcred ASC"
    Set rs = READSET(SQLX$)
    With rs
       Do While Not (.EOF)
         Call TRACE(24, JJ&, FIN&)
         SALDO# = XVALO(!IDEBECOMP)
         If Abs(SALDO#) > 0.005 Then
           DOCORIG$ = SAFETEXT$(!codicom_lar)
           DESCMOV$ = SAFETEXT$(!DESCRIMOV)
           FEMI$ = FETEXCOR1$(CVINT(!FECHEMISI))
           FeVencI$ = FETEXCOR1$(CVINT(!fechacred))
           NCLIEN = SAFETEXT$(!NUCLIEN)
           MON% = XVALO(!Mone_Emis)
           COTIDOL# = XVALO(!COTIZA_DOLAR)
           If COTIDOL# <= 0 Then COTIDOL# = 1
           If CHRECIX% = 0 Then 'SOLO INGRESE UNA VEZ
             DESCRIPCIONFILA$ = "Cheq x Cobrar"
             CUCONTABANCO$ = CUTACBLESEGUNLEYENDA$(DESCRIPCIONFILA$)
             CHRECIX% = 1
           End If
           Call AGREGACASHFLOW(RET_IDCABEZA, NCLIEN, 0, "CHRECI", SALDO#, FEMI$, FeVencI$, DOCORIG$, DESCMOV$, MON%, COTIDOL#, "Cheq x Cobrar", 200, 0, "", CUCONTABANCO$)
         End If
         JJ& = JJ& + 1
         .MoveNext
       Loop
    End With
    rs.Close
    Set rs = Nothing
    'TARJETAS DE CREDITO A COBRAR EN EL RANGO SELECCIONADO
    Condi$ = "TipoMovim = 'CHRECI'  AND FechAcred >='" & DESDEX$ & "' AND FECHACRED <='" & HASTAX$ & "'"

    FIN& = CantRegistros("CAJABANC", Condi$)
    JJ& = 1
    SQLX$ = "SELECT Idebecomp,codicom_lar,DESCRIMOV,FECHEMISI,fechacred,NUCLIEN,Mone_Emis,COTIZA_DOLAR,OPCIMOVIM,CUECONTAI FROM CAJABANC WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE " + Condi$
    SQLX$ = SQLX$ + " AND (OPCIMOVIM = 'TC.' OR OPCIMOVIM = 'TC' OR OPCIMOVIM = 'ME') AND (MARCONSI = '' or MARCONSI is null)"
    SQLX$ = SQLX$ + " ORDER BY FechAcred ASC"
    Set rs = READSET(SQLX$)
    With rs
       Do While Not (.EOF)
         Call TRACE(24, JJ&, FIN&)
         SALDO# = XVALO(!IDEBECOMP)
         If Abs(SALDO#) > 0.005 Then
           DESCMOV$ = SAFETEXT$(!DESCRIMOV)
           COEF = COEFI_IMPOCONCOMISION#(TRIM$(Left$(DESCMOV$, 18)))  ' le descuenta la comision
           SALDO# = SALDO# * COEF
           DOCORIG$ = SAFETEXT$(!codicom_lar)
           FEMI$ = FETEXCOR1$(CVINT(!FECHEMISI))
           FeVencI$ = FETEXCOR1$(CVINT(!fechacred))
           NCLIEN = SAFETEXT$(!NUCLIEN)
           MON% = XVALO(!Mone_Emis)
           COTIDOL# = XVALO(!COTIZA_DOLAR)
           If COTIDOL# <= 0 Then COTIDOL# = 1
           TIPODOCX$ = SAFETEXT$(!OpciMovim)
           If TIPODOCX$ = "TC" Then
              LEYENAMOSTAR$ = "Tarj.a Acred."
           Else
              LEYENAMOSTAR$ = "Medios Electrónicos"
           End If
           CCBLE% = XVALO(!CUECONTAI)
           COCUEX$ = CODCUE$(CCBLE%)
           DESCRIPCIONFILA$ = LEYENAMOSTAR$
           CUCONTABANCO$ = CUTACBLE_TC_ME_SEGUN_CUENTACAJABANC$(COCUEX$)
           Call AGREGACASHFLOW(RET_IDCABEZA, NCLIEN, 0, TIPODOCX$, SALDO#, FEMI$, FeVencI$, DOCORIG$, DESCMOV$, MON%, COTIDOL#, LEYENAMOSTAR$, 220, CCBLE%, COCUEX$, CUCONTABANCO$)
         End If
         JJ& = JJ& + 1
         .MoveNext
       Loop
    End With
    rs.Close
    Set rs = Nothing
    'FACTURAS PENDIENTES DE COBRANZA
    Condi$ = "FEVENCIM <= '" + HASTAX$ + "'"
    FIN& = CantRegistros("SADEUPEN", Condi$)
    JJ& = 1
    FVENX% = 0
    SQLX$ = "SELECT NUCLIEN,FEVENCIM,CodiCom_Lar,FECHEMISI, SALDDEBE AS SUDEB, SALDACRE AS SUHAB,Mone_Emis"
    SQLX$ = SQLX$ + " FROM SADEUPEN WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE " + Condi$
    
    Set rs = READSET(SQLX$)
    With rs
       Do While Not (.EOF)
         Call TRACE(24, JJ&, FIN&)
         DEU# = XVALO(!SUDEB) - XVALO(!SUHAB)
         If Abs(DEU#) > 0.005 Then
           DOCORIG$ = SAFETEXT$(!codicom_lar)
           DESCMOV$ = ""
           FEMI$ = FETEXCOR1$(CVINT(!FECHEMISI))
           FeVencI$ = FETEXCOR1$(CVINT(!FEVENCIM))
           NCLIEN = SAFETEXT$(!NUCLIEN)
           MON% = XVALO(!Mone_Emis)
           COTIDOL# = 1
           If COTIDOL# <= 0 Then COTIDOL# = 1
           DESCRIPCIONFILA$ = "Fact.x Cobrar"
           If FVENX% = 0 Then
            CUCONTABANCO$ = CUTACBLESEGUNLEYENDA$(DESCRIPCIONFILA$) 'PARA QUE ENTRE SOLO UNA VEZ
            FVENX% = 1
           End If
           Call AGREGACASHFLOW(RET_IDCABEZA, NCLIEN, 0, "FVEN", DEU#, FEMI$, FeVencI$, DOCORIG$, DESCMOV$, MON%, COTIDOL#, "Fact.x Cobrar", 100, 0, "", CUCONTABANCO$)
         End If
         JJ& = JJ& + 1
         .MoveNext
       Loop
    End With
    rs.Close
    Set rs = Nothing
    ' CHEQUES A DEBITAR
    If ULTREG&("CUEBANC") > 0 Then
        DESDEMENOS30DIAS% = DIASANTES(DES%, 30)  'SE USA PARA LOS CHEQUES PROPIOS A DEBITAR RETROCEDE EL DESDE 30 DIAS
        DESDEMENOS30DIAS1$ = FETEXCOR1$(DESDEMENOS30DIAS%)
        Condi$ = "TipoMovim = 'CHENTR' AND  (MARCONSI = '' or MARCONSI is null)"
        Condi$ = Condi$ + "AND FechAcred>='" & DESDEMENOS30DIAS1$ & "' AND FECHACRED<='" & HASTAX$ & "' AND ("
        
        For IJ& = 1 To ULTREG&("CUEBANC")
            XX$ = REGLEIDO$("CUEBANC", IJ&)
            CCON% = CUEINT%(Mid$(XX$, 1, 12))
            Condi$ = Condi$ + "CueContaI = " & CCON% & " OR "
        Next IJ&
        Condi$ = Mid(Condi$, 1, Len(Condi$) - 3) + ")"
        FIN& = CantRegistros("CAJABANC", Condi$)
        JJ& = 0
        '
        SQLX$ = "SELECT Idebecomp,codicom_lar,DESCRIMOV,FECHEMISI,FechAcred,NUCLIEN,Mone_Emis,COTIZA_DOLAR,"
        SQLX$ = SQLX$ + " IHABECOMP,NUPROVE,CUECONTAI FROM CAJABANC WITH(NOLOCK)"
        SQLX$ = SQLX$ + "WHERE " + Condi$
        SQLX$ = SQLX$ + " ORDER BY FechAcred ASC;"
        Set rs = READSET(SQLX$)
        With rs
           Do While Not (.EOF)
             Call TRACE(24, JJ&, FIN&)
             FECDEBI% = CVINT(!fechacred)
             SALDO# = XVALO(!IHABECOMP)
             If Abs(SALDO#) > 0.005 Then
                DOCORIG$ = SAFETEXT$(!codicom_lar)
                DESCMOV$ = SAFETEXT$(!DESCRIMOV)
                FEMI$ = FETEXCOR1$(CVINT(!FECHEMISI))
                FeVencI$ = FETEXCOR1$(CVINT(!fechacred))
                NCLIEN = XVALO(!NUCLIEN)
                NUPROV = XVALO(!NUPROVE)
                MON% = XVALO(!Mone_Emis)
                COTIDOL# = XVALO(!COTIZA_DOLAR)
                If COTIDOL# <= 0 Then COTIDOL# = 1
                CCBLE% = XVALO(!CUECONTAI)
                COCUEX$ = CODCUE$(CCBLE%)
                DESCRIPCIONFILA$ = "Cheq x Pagar"
                CUCONTABANCO$ = COCUEX$
                Call AGREGACASHFLOW(RET_IDCABEZA, NCLIEN, NUPROV, "CHENTR", SALDO#, FEMI$, FeVencI$, DOCORIG$, DESCMOV$, MON%, COTIDOL#, "Cheq x Pagar", 1100, CCBLE%, COCUEX$, CUCONTABANCO$)
             End If
             JJ& = JJ& + 1
             .MoveNext
           Loop
        End With
        rs.Close
        Set rs = Nothing
    End If
    '
    'FACTURAS PENDIENTES DE PAGO
    Condi$ = "FEVENCIM<='" & HASTAX$ & "'"
    FIN& = CantRegistros("SACREPEN", Condi$)
    JJ& = 0
    '
    FCOMX% = 0
    SQLX$ = "SELECT NUCLIEN,NUPROVE,FEVENCIM,CODICOM_COR,CodiCom_Lar,FECHEMISI, SALDDEBE AS SUDEB, SALDACRE AS SUHAB, Mone_Emis"
    SQLX$ = SQLX$ + " FROM SACREPEN WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE " + Condi$
    Set rs = READSET(SQLX$)
    With rs
       Do While .EOF = False
          Call TRACE(24, JJ&, FIN&)
          SALDO# = XVALO(!SUHAB) - XVALO(!SUDEB)
          If Abs(SALDO#) > 0.005 Then
                DOCORIG$ = SAFETEXT$(!codicom_lar)
                DESCMOV$ = ""
                FEMI$ = FETEXCOR1$(CVINT(!FECHEMISI))
                FeVencI$ = FETEXCOR1$(CVINT(!FEVENCIM))
                NCLIEN = XVALO(!NUCLIEN)
                NUPROV = XVALO(!NUPROVE)
                MON% = XVALO(!Mone_Emis)
                COTIDOL# = 1
                DESCRIPCIONFILA$ = "Compras"
                If FCOMX% = 0 Then
                   CUCONTABANCO$ = CUTACBLESEGUNLEYENDA$(DESCRIPCIONFILA$)
                   FCOMX% = 1
                End If


                Call AGREGACASHFLOW(RET_IDCABEZA, NCLIEN, NUPROV, "FCOM", SALDO#, FEMI$, FeVencI$, DOCORIG$, DESCMOV$, MON%, COTIDOL#, "Compras", 1000, 0, "", CUCONTABANCO$)
          End If
          JJ& = JJ& + 1
       .MoveNext
       Loop
    End With
    rs.Close
    Set rs = Nothing
    '
    IDL = XVALO(Control("", "CTALCUBA"))
    DESC$ = Control("", "DENOCUBA")
    If IDL > 0 Then
        J& = 0
        SQLX$ = "SELECT CAMPO1, CAMPO2,CAMPO3 FROM DATASQL WITH(NOLOCK) WHERE ID_LISTA = " & IDL & " And LISTA = '" & DESC$ & "'"
        SQLX$ = SQLX$ + " order by campo2"
        Set rs = READSET(SQLX$)
        With rs
          Do While Not .EOF
            CTACBLE$ = SAFETEXT$(!CAMPO1)
            CUCONTABANCO$ = CTACBLE$
            CCBLE% = CUEINT%(CTACBLE$)
            DESCRBANCO$ = SAFETEXT$(!CAMPO2)
            SALDO# = XVALO(!CAMPO3)
            Call AGREGACASHFLOW(RET_IDCABEZA, 0, 0, "SALBANCO", SALDO#, DESDEX$, DESDEX$, DESCRBANCO$, CTACBLE$ & " " & DESCRBANCO$, 1, 1, "Saldo Banco", 80, CCBLE%, COCUEX$, CUCONTABANCO$)
            .MoveNext
          Loop
        End With
    End If
    IDCABEZA = 0
ERROR_GUARDAR:
    If Err <> 0 Then
       MsgBox Err.Description
       FLEXCONN.RollbackTrans
       On Error GoTo 0
       Screen.MousePointer = 1

       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION")
       Exit Sub
    Else
      'SE GUARDA EL CONTEXTO EN LAS TABLAS
      FLEXCONN.CommitTrans
      Call COMUNI("Grabación Realizado Exitosamente")
      IDCABEZA = RET_IDCABEZA
      On Error GoTo 0
    End If
    
    If IDCABEZA < 1 Then
       Screen.MousePointer = 1
       Call COMUNI("No se ha Realizado Correctamente la Generación del Cash-Flow\Vuelva a Intentar")
       Exit Sub
    End If
    Screen.MousePointer = 1

    Call AGREGARVALORESCASHFLOW(PRESFLOW.MSF1, MODX$, IDCABEZA)



'consulta formato mensual
'select year(fechven) as ano, month(fechven) as mes,sum(importe) as importex,tipodocu   from backupflow where id_cabezera = 8
'group by   year(fechven) , month(fechven) ,tipodocu
'order by year(fechven)  , month(fechven) ,tipodocu

'consulta formato semanal
'select year(fechven) as ano, month(fechven) as mes,DATEPART(wk, fechven) as semana,sum(importe) as importex,tipodocu   from backupflow where id_cabezera = 8
'group by   year(fechven) , month(fechven) ,DATEPART(wk, fechven),tipodocu
'order by year(fechven)  , month(fechven) ,DATEPART(wk, fechven), tipodocu
 
'consulta formato diario
'select year(fechven) as ano, month(fechven) as mes,day(fechven) as dia ,sum(importe) as importex,tipodocu   from backupflow where id_cabezera = 8
'group by   year(fechven) , month(fechven) ,day(fechven),tipodocu
'order by year(fechven)  , month(fechven) ,day(fechven), tipodocu
 
 'para obtener el numero de semana pero dentro del mes.
 
' select datepart(week,'17/JUN/2019') -
'datepart(week, dateadd(dd,-day('17/JUN/2019')+1,'17/JUN/2019')) +1
99
End Sub

Function CUTACBLE_TC_ME_SEGUN_CUENTACAJABANC$(COCUEX$)
   Condi$ = "LISTA = 'CONFIMECO' AND CAMPO1 = '" & COCUEX$ & "'"
   CUTACBLE_TC_ME_SEGUN_CUENTACAJABANC$ = VALOBADA("DATASQL", "CAMPO5", Condi$, "NOMESS")
End Function
 Sub AGREGARVALORESCASHFLOW(MSF As MSFlexGrid, TIPOPERIODO$, IDKABEZA)
    'ARMADO DE CONDICION DEPENDIENDO EL TIPO DE PERIODO
'    PARA DIAS BIEN
'    HAY QUE COMPLETAR LOS NOMBRES DE LAS FILAS
    'IMPUESTOS
    IDL = XVALO(Control("", "CTAREPER"))
    DESC$ = Control("", "DENREPER")
    CANTIMPUESTOS% = CantRegistros&("DATASQL", " ID_LISTA = " & IDL & " And LISTA = '" & DESC$ & "'", "NOMESS", , "CAMPO2")
    If CANTIMPUESTOS% > 0 Then
       J& = 0
       Dim DESCRIMPU$()
       SQLX$ = "SELECT  DISTINCT CAMPO2,CAMPO3 FROM DATASQL WITH(NOLOCK) WHERE  ID_LISTA = " & IDL & " And LISTA = '" & DESC$ & "'"
       SQLX$ = SQLX$ + " order by campo2"
        Set rs = READSET(SQLX$)
        With rs
          Do While Not .EOF
            J& = J& + 1
            ReDim Preserve DESCRIMPU$(J&)
            DESCRIMPU$(J&) = SAFETEXT$(!CAMPO3)
            .MoveNext
          Loop
        End With
        rs.Close
        Set rs = Nothing
    End If
    ''''''''''''''''''''''''''''''''''''''''''''
    IDL2 = XVALO(Control("", "CTAINGGAS"))
    DESC2$ = TRIM$(Control("", "DENINGGAS"))
    CANTINGRGAS% = CantRegistros&("DATASQL", " ID_LISTA = " & IDL2 & " And LISTA = '" & DESC2$ & "'", "NOMESS", , "CAMPO2")
    If CANTINGRGAS% > 0 Then
        SQLX$ = "SELECT CAMPO1, CAMPO2,CAMPO3,CAMPO4,CAMPO5 FROM DATASQL WITH(NOLOCK) WHERE ID_LISTA = " & IDL2 & " And LISTA = '" & DESC2$ & "'"
        SQLX$ = SQLX$ + " ORDER BY CAMPO1"
        Dim DESCRIINGRGAS$()
        
        Set rs = READSET(SQLX$)
        With rs
          Do While Not .EOF
            J& = J& + 1
            ReDim Preserve DESCRIMPU$(J&)
            DESCRIMPU$(J&) = SAFETEXT$(!CAMPO3)
            .MoveNext
          Loop
        End With
        rs.Close
        Set rs = Nothing
    End If
    ''''''''''''''''''''''''''''''''''''''''''''
    For i& = 1 To MSF.Cols - 1
        LEIDO$ = MSF.TextMatrix(0, i&)
        If TIPOPERIODO$ = "DIA" Then
           FE% = FECHANUM(LEIDO$)
           If FE% < 1 Then GoTo 70
           If i& = 1 Then
              Condi$ = "ACREDITACION <= '" & FETEXCOR1$(FE%) & "'"
           Else
              Condi$ = "ACREDITACION = '" & FETEXCOR1$(FE%) & "'"
           End If
        ElseIf TIPOPERIODO$ = "SEM" Then
           FEIN% = FECHANUM(Left(LEIDO$, 8))
           FEFI% = FECHANUM(Mid$(LEIDO$, 10, 8))
           If FEIN% < 1 Then GoTo 70
           XX = RangoSemanaSegunFecha$(FEIN%, RET1, RET_NUMSEMANA%) 'DE ESTA RUTINA TOMO EL NUMERO DE SEMANA
           NUSEM% = RET_NUMSEMANA% 'LO RETORNA EN EL PARAMETRO RET_NUMSEMANA%
           If i& = 1 Then
              'CONDI$ = " DATEPART(wk, ACREDITACION) <= " & NUSEM%
              Condi$ = "ACREDITACION <= '" & FETEXCOR1$(FEFI%) & "'"
           Else
              'CONDI$ = " DATEPART(wk, ACREDITACION) = " & NUSEM%
              Condi$ = "ACREDITACION >= '" & FETEXCOR1$(FEIN%) & "' AND ACREDITACION <= '" & FETEXCOR1$(FEFI%) & "' "
           End If
        ElseIf TIPOPERIODO$ = "MES" Then
           NMES% = NUMMES_TRESCARACTERES%(Left(LEIDO$, 3))
           If NUMES% < 1 Then GoTo 70
           ANO% = XVALO(Mid$(LEIDO$, 6, 4))
           If i& = 1 Then
              Condi$ = " YEAR(ACREDITACION) <= " & ANO%
              Condi$ = Condi$ + " AND  MONTH(ACREDITACION) <= " & NMES%
           Else
              Condi$ = " YEAR(ACREDITACION) = " & ANO%
              Condi$ = Condi$ + " AND  MONTH(ACREDITACION) = " & NMES%
           End If
        End If
        '

        Dim IMPUESTOSCONINFO()
        ReDim Preserve IMPUESTOSCONINFO(CANTIMPUESTOS% + CANTINGRGAS%)
        
          'Call CARGARFILA(i&, CONDI$, MSF)
           SQLX$ = "select sum(importe) as importex, tipodocu, ORDENAMIENTO  from backupflow WITH(NOLOCK)"
           SQLX$ = SQLX$ + " where id_cabezera = " & IDKABEZA & " AND " & Condi$
           SQLX$ = SQLX$ + " GROUP BY TIPODOCU,ORDENAMIENTO"
           Set rs = READSET(SQLX$)
           With rs
             Do While Not .EOF
               DESCRIPCIONFILA$ = ""
               TIDOC$ = SAFETEXT$(!TIPODOCU)
               Select Case TIDOC$
                  Case "FVEN"
                     FILA% = 1
                     DESCRIPCIONFILA$ = "Fact.x Cobrar"
                  Case "FCOM"
                     FILA% = 2
                     DESCRIPCIONFILA$ = "Compras"
                  Case "CHENTR"
                     FILA% = 3
                     DESCRIPCIONFILA$ = "Cheq x Pagar"
                  Case "CHRECI"
                     FILA% = 4
                     DESCRIPCIONFILA$ = "Cheq x Cobrar"
                  Case "TC"
                     FILA% = 5
                     DESCRIPCIONFILA$ = "Tarj.a Acred."
                  Case "SALBANCO"
                     FILA% = 6
                     DESCRIPCIONFILA$ = "Saldo Banco"
                  Case "ME"
                     FILA% = 7
                     DESCRIPCIONFILA$ = "Medios Electrónicos"
                  Case Else
                  
                     'ACA BUSCA LA COINCIDENCIA POR EL TIPO DE DOCUMENTO YA SEA DE LOS IMPUESTOS O DE GASTOS
                     If CANTIMPUESTOS% > 0 Or CANTINGRGAS% > 0 Then
                        For L& = 1 To J&
                           If TIDOC$ = DESCRIMPU$(L&) Then
                              FILA% = 7 + L&
                              Exit For
                           End If
                        Next L&
                     End If
               End Select
               If FILA% >= MSF.Rows Then MSF.Rows = FILA% + 1
               MSF.TextMatrix(FILA%, i&) = FORMATNUM$(XVALO(!importex), "F14.2")
               MSF.TextMatrix(FILA%, MSF.Cols - 1) = XVALO(!Ordenamiento)
'               MSF.TextMatrix(FILA%, MSF.Cols - 1) = DESCRIPCIONFILA$
               .MoveNext
             Loop
           End With
70    Next i&
'ORDENAMIENTO DEL MSFFLEXGRID
    NOMBREFILAS$ = "Fact.x Cobrar;Compras;Cheq x Pagar;Cheq x Cobrar;Tarj.a Acred.;Saldo Banco;Medios Electrónicos"
    If CANTIMPUESTOS% > 0 Or CANTINGRGAS% > 0 Then
        For r& = 1 To UBound(DESCRIMPU$)
'           If IMPUESTOSCONINFO(R&) = 1 Then
            NOMBREFILAS$ = NOMBREFILAS$ & ";" & DESCRIMPU$(r&)
'           End If
        Next r&
    End If
    
    NOMBREFILAS$ = NOMBREFILAS$ & ";" & "DETALLE DE INGRESOS"
    NOMBREFILAS$ = NOMBREFILAS$ & ";" & "DETALLE DE EGRESOS"
    NOMBREFILAS$ = NOMBREFILAS$ & ";" & "TOTAL INGRESOS"
    NOMBREFILAS$ = NOMBREFILAS$ & ";" & "TOTAL EGRESOS"
    
    Call AGREGAR_NombreFilas(PRESFLOW.MSF1, NOMBREFILAS$, 1, 0)
    MSF.COL = MSF.Cols - 1 'xx Columan por la q quieres ordenar
    MSF.Sort = 3 'Aplicar orden al grid
    For i& = 1 To MSF.Rows - 1
       If MSF.TextMatrix(i&, 0) = "DETALLE DE INGRESOS" Then
          Call COLOREAR_GRILLA_SOLO_UNA_AFILA(MSF, &HFF00&, i&)
          KOLOR = &HC0FFC0
          GoTo 80
       End If

       If MSF.TextMatrix(i&, 0) = "TOTAL INGRESOS" Then
          Call COLOREAR_GRILLA_SOLO_UNA_AFILA(MSF, &HFF00&, i&)
          KOLOR = &HFFFFFF
          GoTo 80
       End If
       If MSF.TextMatrix(i&, 0) = "DETALLE DE EGRESOS" Then
          Call COLOREAR_GRILLA_SOLO_UNA_AFILA(MSF, &HFFFF&, i&)
          KOLOR = &HC0FFFF
          GoTo 80
       End If
       If MSF.TextMatrix(i&, 0) = "TOTAL EGRESOS" Then
          Call COLOREAR_GRILLA_SOLO_UNA_AFILA(MSF, &HFFFF&, i&)
          KOLOR = &HFFFFFF
          GoTo 80
       End If

       Call COLOREAR_GRILLA_SOLO_UNA_AFILA(MSF, KOLOR, i&)
80    Next i&
      Call TOTALIZARCASHFLOW(PRESFLOW.MSF1)
      
      PRESFLOW.TXTMODO = TIPOPERIODO$
      PRESFLOW.TXTID = IDKABEZA
      PRESFLOW.Show 1
      
 End Sub
Sub AGREGACASHFLOW(IDCABEZA, NCLIEN, NUPROV, TIPODMOV$, SALDO#, FEMI$, FeVenc$, DOCORIG$, DESCMOV$, MON%, COTIDOL#, TIPOMOVIMAMOSTRAR$, CODIGOORDEN&, Optional CODINCUEX%, Optional CODEXTCUE$, Optional CUCONTABANCO$)
        FEACREDITACION$ = FeVenc$
        fev% = CVINT(FeVenc$)
        DIAX% = 0
        
        
        
        'CALCULO LA DEMORA CONFIGURADA DE ACREDITACION SOLO PARA CHEQUES A COBRAR Y FACTURAS DE VENTA
        If TIPODMOV$ = "CHRECI" Then
           DIAX% = XVALO(Control("", "CLEABANC")) ' SI TIENE CANTIDAD DE DIAS CALCULA QUE CAIGA EN UN DIA HABIL
           If DIAX% > 0 Then
             fev% = fev% + CANT_DIASPOSTHABILES(fev%, DIAX%)
           End If

        ElseIf TIPODMOV$ = "FVEN" Then
           DIAX% = XVALO(Control("", "EFECFACT"))
           If DIAX% > 0 Then
             fev% = fev% + CANT_DIASPOSTHABILES(fev%, DIAX%) ' SI TIENE CANTIDAD DE DIAS CALCULA QUE CAIGA EN UN DIA HABIL
           End If
        End If
        '
        'ESTO ES PARA TODOS LOS CASOS SI LA FECHA DE VENCIMIENTO ES SAB O DOMINGO LO PONE EN LUNES.
        If fev% > 33 Then
         If DIASEM%(fev%) = 0 Then fev% = DIASPOST(fev%, 1)
         If DIASEM%(fev%) = 6 Then fev% = DIASPOST(fev%, 2)
         FEACREDITACION$ = FETEXCOR1$(fev%)
        End If
        
        If MON% = 1 And Not (COTIDOL# > 0.005) Then
            COTIDOL# = 1
        End If

        SQLX$ = "INSERT INTO BACKUPFLOW "
        SQLX$ = SQLX$ + " (ID_CABEZERA,NUCLIEN,NUPROVE,TIPODOCU,IMPORTE,FECHEMISI,FECHVEN,DOCUORIGEN,"
        SQLX$ = SQLX$ + " DescriMov,Mone_Emis,COTIZA_DOLAR,ACREDITACION,MOVCASHFLOW,ORDENAMIENTO,CodInCue,"
        SQLX$ = SQLX$ + " CodExCue,CUCONTA_BANCO)"
        
        SQLX$ = SQLX$ + " VALUES(" & IDCABEZA
        SQLX$ = SQLX$ + " ," & XVALO(NCLIEN)
        SQLX$ = SQLX$ + " ," & XVALO(NUPROV)
        SQLX$ = SQLX$ + " ,'" & TIPODMOV$ & "'"
        SQLX$ = SQLX$ + " ,'" & FORMATNUM$(SALDO#, "F16.5") & "'"
        SQLX$ = SQLX$ + " ,'" & FEMI$ & "'"
        SQLX$ = SQLX$ + " ,'" & FeVenc$ & "'"
        SQLX$ = SQLX$ + " ,'" & DOCORIG$ & "'"
        SQLX$ = SQLX$ + " ,'" & DESCMOV$ & "'"
        SQLX$ = SQLX$ + " ," & MON%
        SQLX$ = SQLX$ + " ," & TRIM$(FORMATNUM$(XVALO(COTIDOL#), "F16.5"))
        SQLX$ = SQLX$ + " ,'" & FEACREDITACION$ & "'"
        SQLX$ = SQLX$ + " ,'" & TIPOMOVIMAMOSTRAR$ & "'"
        SQLX$ = SQLX$ + " ," & CODIGOORDEN&
        SQLX$ = SQLX$ + " ," & CODINCUEX%
        SQLX$ = SQLX$ + " ,'" & Left$(CODEXTCUE$, 16) & "'"
        SQLX$ = SQLX$ + " ,'" & Left$(CUCONTABANCO$, 16) & "'" 'CUENTA CONTABLE PARA ARMAR PRESENTACION POSICION DE BANCO
        
        SQLX$ = SQLX$ + ")"

        FLEXCONN.Execute (SQLX$)

End Sub

Private Sub Command4_Click()
    '
    Command4.Enabled = False
    
    If InStr(1, EMPREAC$, "DOSIVAC") > 0 Then
5       OPX% = ALTERNA%("Egresos e Ingresos Varios\Sueldos\Cargas Sociales\Interbankin\Gastos Estructura")
        If OPX% < 1 Or OPX% > 5 Then GoTo 99
        '
        Call GRABAINGREGAS(OPX%)
        GoTo 5
    End If

    Call APERBASDAT("CABAN")
    Call SETULTREG("PLAGASPE", 0)
    J& = 0
    SQLX$ = "SELECT * FROM PLAGASPEN WITH(NOLOCK)"
    Dim PLAGTM As ADODB.Recordset
    Set PLAGTM = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With PLAGTM
       Do While Not .EOF
          '
          XX$ = REGBLAN("PLAGASPE")
          FE% = CVINT(!FECHA)
          Call REPLA(XX$, MKI$(FE%), 1, 2)
          Call REPLA(XX$, (SAFETEXT$(!REFERENCIA)), 3, 64)
          Call REPLA(XX$, MKD(XVALO(!INGRESOS)), 67, 8)
          Call REPLA(XX$, MKD(XVALO(!GASTOS)), 75, 8)
          J& = J& + 1
          Call GRAREG("!PLAGASPE", XX$, J&)
          .MoveNext   '
       Loop
    End With
    Call SETULTREG&("PLAGASPE", J&)
    PLAGTM.Close
    Set PLAGTM = Nothing
    
10  VTB% = VENTABU%("PLAGASPE/NC/TITUPAN=Plan de Gastos e Ingresos")
    If VTB% < 0 Then GoTo 99
     'LEER LA VENTANA Y VALIDAR
    For U& = 1 To ULTREG&("!PLAGASPE")
         x$ = REGLEIDO$("!PLAGASPE", U&)
         FE% = CVI(Mid$(x$, 1, 2))
         REF$ = Mid(x$, 3, 64)
         ING# = CVD(Mid$(x$, 67, 8))
         GAST# = CVD(Mid$(x$, 75, 8))
         '
         If XVALO(FE%) < 1 Then
           Call COMUNI("Falta Ingresar Fecha en Fila: " & U&)
           GoTo 10
         End If
         '
         If TRIM$(REF$) = "" Then
           Call COMUNI("Falta Ingresar Referencia en Fila: " & U&)
           GoTo 10
         End If
         '
         If XVALO(ING#) < 0.005 And XVALO(GAST#) < 0.005 Then
           Call COMUNI("Falta Ingresar Importe en Fila: " & U&)
           GoTo 10
         End If
         '
         If XVALO(ING#) > 0.005 And XVALO(GAST#) > 0.005 Then
           Call COMUNI("No Válido Ingresar 2 Importes en Fila: " & U&)
           GoTo 10
         End If

    Next U&
     '
    Call SETULTREG&("PLAGASPE", U&)
    
    Screen.MousePointer = 11
    '
    ' SI PASO LA VALIDACION MARBORRA = 1
    
    Condi$ = "REFERENCIA <> ''"
    Call ACTUREGISTRO("PLAGASPEN", "MARBORRA", Condi$, 1, REGAF&, "NOMESS")
    
    ' RECORRER LA VENTANA Y GRABAR EN B.DATOS
    SQLX$ = "SELECT * FROM PLAGASPEN"
    SQLX$ = SQLX$ + " WHERE REFERENCIA = '1'"
    Dim PLAGTMP As ADODB.Recordset
    For U& = 1 To ULTREG&("!PLAGASPE")
        Set PLAGTMP = New ADODB.Recordset
        PLAGTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        With PLAGTMP
           .AddNew   ' YA CONVERTIDO
           On Error Resume Next
           !CODIEMPR = CodiEmp%
           On Error GoTo 0
           x$ = REGLEIDO$("!PLAGASPE", U&)
           FE% = CVI(Mid$(x$, 1, 2))
           REF$ = Mid(x$, 3, 64)
           ING# = CVD(Mid$(x$, 67, 8))
           GAST# = CVD(Mid$(x$, 75, 8))
           '
           !FECHA = CVDAT(FE%)
           !REFERENCIA = REF$
           !INGRESOS = ING#
           !GASTOS = GAST
           !MARBORRA = 0
           .Update
        End With
        PLAGTMP.Close
        Set PLAGTMP = Nothing
    Next U&
    '
    Call BORRAREGISTROS("PLAGASPEN", "MARBORRA=" & 1, REG&)
    Screen.MousePointer = 1
    '
99  Command4.Enabled = True
    '
End Sub

Sub GRABAINGREGAS(MODO%)
    '
    Call APERBASDAT("CABAN")
    Call PoneEnCeroNull("PLAGASPEN", "TIPO", REGAF&)
    Call SETULTREG("!PLAGASPE", 0)
    J& = 0
    Tipo% = MODO% - 1
    SQLX$ = "SELECT * FROM PLAGASPEN WITH(NOLOCK) "
    SQLX$ = SQLX$ & " WHERE TIPO= " & Tipo%
    Dim PLAGTM As ADODB.Recordset
    Set PLAGTM = READSET(SQLX$)
    With PLAGTM
       Do While Not .EOF
          '
          XX$ = REGBLAN("PLAGASPE")
          FE% = CVINT(!FECHA)
          Call REPLA(XX$, MKI$(FE%), 1, 2)
          Call REPLA(XX$, (SAFETEXT$(!REFERENCIA)), 3, 64)
          Call REPLA(XX$, MKD(XVALO(!INGRESOS)), 67, 8)
          Call REPLA(XX$, MKD(XVALO(!GASTOS)), 75, 8)
          J& = J& + 1
          Call GRAREG("!PLAGASPE", XX$, J&)
          .MoveNext   '
       Loop
    End With
    Call SETULTREG&("!PLAGASPE", J&)
    PLAGTM.Close
    Set PLAGTM = Nothing
    
    Select Case MODO%
        Case 1: TITU$ = "Plan de Gastos e Ingresos"
        Case 2: TITU$ = "Plan de Sueldos y Jornales"
        Case 3: TITU$ = "Plan de Cargas Sociales"
        Case 4: TITU$ = "Plan de Interbankin"
        Case 5: TITU$ = "Plan de Gastos de Estructura"
        Case 30: TITU$ = "Tabla de Bancos"
    End Select
    
    ATRI$ = "/POSEDIT(3)=2"
    If MODO% > 25 Then ATRI$ = "/POSEDIT(4)=2"
10  VTB% = VENTABU%("PLAGASPE/NC/TITUPAN=" & TITU$ & ATRI$)
    If VTB% < 0 Then GoTo 99
     'LEER LA VENTANA Y VALIDAR
    For U& = 1 To ULTREG&("!PLAGASPE")
         x$ = REGLEIDO$("!PLAGASPE", U&)
         FE% = CVI(Mid$(x$, 1, 2))
         REF$ = Mid(x$, 3, 64)
         ING# = CVD(Mid$(x$, 67, 8))
         GAST# = CVD(Mid$(x$, 75, 8))
         '
         If XVALO(FE%) < 1 Then
           Call COMUNI("Falta Ingresar Fecha en Fila: " & U&)
           GoTo 10
         End If
         '
         If TRIM$(REF$) = "" Then
           Call COMUNI("Falta Ingresar Referencia en Fila: " & U&)
           GoTo 10
         End If
         '
         If XVALO(ING#) < 0.005 And XVALO(GAST#) < 0.005 Then
           Call COMUNI("Falta Ingresar Importe en Fila: " & U&)
           GoTo 10
         End If
         '
         If XVALO(ING#) > 0.005 And XVALO(GAST#) > 0.005 Then
           Call COMUNI("No Válido Ingresar 2 Importes en Fila: " & U&)
           GoTo 10
         End If

    Next U&
     '
    Call SETULTREG&("PLAGASPE", U&)
    
    Screen.MousePointer = 11
    '
    ' SI PASO LA VALIDACION MARBORRA = 1
    
    Condi$ = "TIPO = " & Tipo%
    Call ACTUREGISTRO("PLAGASPEN", "MARBORRA", Condi$, 1, REGAF&, "NOMESS")
    
    ' RECORRER LA VENTANA Y GRABAR EN B.DATOS
    SQLX$ = "SELECT * FROM PLAGASPEN"
    SQLX$ = SQLX$ + " WHERE REFERENCIA = '1'"
    Dim PLAGTMP As ADODB.Recordset
    For U& = 1 To ULTREG&("!PLAGASPE")
        Set PLAGTMP = New ADODB.Recordset
        PLAGTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        With PLAGTMP
           .AddNew   ' YA CONVERTIDO
           On Error Resume Next
           !CODIEMPR = CodiEmp%
           On Error GoTo 0
           x$ = REGLEIDO$("!PLAGASPE", U&)
           FE% = CVI(Mid$(x$, 1, 2))
           REF$ = Mid(x$, 3, 64)
           ING# = CVD(Mid$(x$, 67, 8))
           GAST# = CVD(Mid$(x$, 75, 8))
           '
           !FECHA = CVDAT(FE%)
           !REFERENCIA = REF$
           !INGRESOS = ING#
           !GASTOS = GAST
           !Tipo = Tipo%
           !MARBORRA = 0
           .Update
        End With
        PLAGTMP.Close
        Set PLAGTMP = Nothing
    Next U&
    '
    Call BORRAREGISTROS("PLAGASPEN", "MARBORRA=" & 1, REG&)
99  Screen.MousePointer = 1
End Sub


Private Sub Command6_Click()
    Command6.Enabled = False
    
    If InStr(1, EMPREAC$, "DOSIVAC") > 0 Then
10      OPX% = ALTERNA%("Cuentas de Caja\Cuentas de Bancos")
        If OPX% < 1 Or OPX% > 2 Then GoTo 99
        '
        If OPX% = 1 Then
            On Error Resume Next
            Call OPCLAV07.Command6_Click
            If Err Then
               On Error GoTo 0
               Call MENSERR(24, "Imposible Procesar la Petición.\Falta Formulario 'OPCLAV07.FRM' en Proyecto '" + App.EXEName + "'")
            End If
            On Error GoTo 0
        ElseIf OPX% = 2 Then
            Call GRABAINGREGAS(30)
        End If
        GoTo 10
    End If
    
    On Error Resume Next
    Call OPCLAV07.Command6_Click
    If Err Then
       On Error GoTo 0
       Call MENSERR(24, "Imposible Procesar la Petición.\Falta Formulario 'OPCLAV07.FRM' en Proyecto '" + App.EXEName + "'")
    End If
    On Error GoTo 0
    
99  Command6.Enabled = True
End Sub

Private Sub Form_Load()
    '
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me, Picture1)
    COMESA$ = FETEXCOR1$(COMESACT%)
    Condi$ = "FECHA < '" & COMESA$ & "'"
    QTR& = CantRegistros("PLAGASPEN", Condi$)
    If QTR& > 0 Then
       TTXX$ = "Existen Ingresos y Gastos Proyectados\Anteriores al Comienzo del Mes Actual.\   \Elija Modalidad de Procesamiento:\(Si Elige Borrarlos, no Podrá Revertirse)\\Conservarlos\Eliminarlos"
       If COMALTER%(TTXX$) = 2 Then
          Call BORRAREGISTROS("PLAGASPEN", Condi$, RAFE&, "NOMESS")
       End If
    End If
    '
    If InStr(EMPREAC$, "DOSIVAC") > 0 Then
        For i& = 1 To Check1.Count
           Check1(i& - 1).Enabled = False
        Next i&
    End If
End Sub


Private Sub mnuConfigFlujoCaja_Click()
  If DERACCE%("ModiFluc", "Configuraciones Especiales de CashFlow") < 2 Then Exit Sub
  FSCASHFLOW.Show 1
  
End Sub
Sub VERCASHEMITIDOS()
    If DERACCE%("VERCASHF", "Visualizar Cash-Flow Generados") < 2 Then Exit Sub
    
    campos$ = "Nro. /F12;Referencia/A64;Fecha y Hora/A20,FDES/A0;FHAS/A0;INTVLO/A0"
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, campos$, FRMZELECHO)
    FRMZELECHO.Caption = "Cashflow Emitidos"

    Dim obj As New RECORXET
    Set RS1 = obj.RS_CashFlowEmitidos
    Call Carga_MSF_Recordset(RS1, campos$, FRMZELECHO.MSF2, 1)
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    FRMZELECHO.Width = 12000
    FRMZELECHO.Show 1
    If XVALO(RESP_ZELE_VECT(1)) > 0 Then
      ID_CABEZA& = XVALO(RESP_ZELE_VECT(1))
      REFE$ = RESP_ZELE_VECT(2)
      REFEFECHA$ = RESP_ZELE_VECT(3)
      FEDEX$ = RESP_ZELE_VECT(4)
      FEHAS$ = RESP_ZELE_VECT(5)
      INTVLO$ = RESP_ZELE_VECT(6)
      '
      DES% = CVINT(FEDEX$)
      HAS% = CVINT(FEHAS$)
      If DES% <= 0 Or HAS% <= 0 Then
         Call COMUNI("Imposible Mostrar CashFlow Seleccionado.\Datos de Rango Incompletos")
         Exit Sub
      End If
     '
     'DE ACUERDO A QUE INTERVALO SE SELECCIONO CARGA EL ENCABEZADO.
     Select Case INTVLO$
     Case "DIA"
        Call CARGA_ENCABEZADO_DIAS(PRESFLOW.MSF1, PRESFLOW, DES%, HAS%, "F12.2", 1, RETCANTIDAS%, RETNUMPRIMERDIA%, RETNUMULTIMDIA%, NOMBREFILAS$)
     Case "SEM"
        Call CARGA_ENCABEZADO_SEMANA(PRESFLOW.MSF1, PRESFLOW, DES%, HAS%, "F21.2", 1, RETCANTISEMANAS%, RETNUMPRIMERASEMANA%, RETNUMULTIMSEMANA%, NOMBREFILAS$)
     Case "MES"
        Call CARGA_ENCABEZADO_MESES(PRESFLOW.MSF1, PRESFLOW, DES%, HAS%, "F18.2", 1, RETCANTIMESES%, RETNUMPRIMERAMES%, RETNUMULTIMES%, NOMBREFILAS$)
     End Select
     PRESFLOW.Caption = REFE$
     Call AGREGARVALORESCASHFLOW(PRESFLOW.MSF1, INTVLO$, ID_CABEZA&)

    End If

End Sub

