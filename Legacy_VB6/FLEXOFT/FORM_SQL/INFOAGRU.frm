VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form INFOAGRU 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informe Agrupado"
   ClientHeight    =   9225
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   14430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9225
   ScaleWidth      =   14430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   11520
      Top             =   7680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtDescripcionCuenta 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Roboto Mono Medium"
         Size            =   8.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   11400
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   5950
      Width           =   2775
   End
   Begin VB.TextBox txtCuentaContable 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Roboto Mono Medium"
         Size            =   8.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11400
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   5450
      Width           =   2775
   End
   Begin MSFlexGridLib.MSFlexGrid mfgInforme 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   7646
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Roboto Mono Medium"
         Size            =   9
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid mfgDetallePorCuenta 
      Height          =   3495
      Left            =   120
      TabIndex        =   1
      Top             =   5160
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   6165
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Roboto Mono Medium"
         Size            =   9
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblPeriodo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10200
      TabIndex        =   8
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label lblNombreInforme 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   9855
   End
   Begin VB.Label Label1 
      Caption         =   "Cuenta Contable"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11400
      TabIndex        =   4
      Top             =   5160
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Roboto Mono"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7800
      TabIndex        =   3
      Top             =   8920
      Width           =   735
   End
   Begin VB.Label lblTotalCuenta 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Roboto Mono Medium"
         Size            =   9
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   8685
      TabIndex        =   2
      Top             =   8760
      Width           =   2055
   End
   Begin VB.Menu mnuConsultas 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu mnuImpresionInforme 
         Caption         =   "Impresión Informe Agrupado"
      End
      Begin VB.Menu mnuImpresionDetalleCuenta 
         Caption         =   "Impresión Detalle por Cuenta"
      End
   End
End
Attribute VB_Name = "INFOAGRU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lblDesdeFecha As Integer
Public lblHastaFecha As Integer

Private Sub mfgInforme_DblClick()

        FILA& = mfgInforme.Row
        
        CuentaContable% = XVALO(mfgInforme.TextMatrix(FILA&, mfgInforme.Cols - 1))
        If CuentaContable% > 0 Then
        
            DesdeFecha% = CInt(lblDesdeFecha)
            HastaFecha% = CInt(lblHastaFecha)
            txtCuentaContable = CODCUE(CuentaContable%)
            txtDescripcionCuenta = DENOCUE(CuentaContable%)
            
            CAMPOS$ = "Fecha/D8;Documento/A18;Referencia/A52;Importe/F12.2;FechaEEUU/A0"
            Call CARGA_ENCABEZADO(mfgDetallePorCuenta, CAMPOS$, FRMZELECHO, , , 0)
            
            ConsultaSql$ = "select fechconta, codicom_lar, refecom, idebecomp, ihabecomp from cajabanc with(nolock) "
            ConsultaSql$ = ConsultaSql$ + " where (FechConta  >= '" & FETEXCOR1$(DesdeFecha%) & "' and fechconta <= '" & FETEXCOR1$(HastaFecha%, "HASTA") & "') "
            ConsultaSql$ = ConsultaSql$ + " and cuecontai=" & CStr(CuentaContable%)
            Dim CABTEMP As ADODB.Recordset
            Set CABTEMP = READSET(ConsultaSql$)
            '
            mfgDetallePorCuenta.Rows = CABTEMP.RecordCount + 1
            JJ& = 0
            With CABTEMP
                 Do While Not .EOF
                    
                    FECHA$ = FECHATEX$(CVINT(!fechconta))
                    DOCUMENTO$ = !codicom_lar
                    REFERENCIA$ = !refecom
                    IMPORTE$ = FORMATNUM$(!idebecomp - !ihabecomp, "F12.2")
                        FECHA0$ = FECHATEXLAR$(CVINT(!fechconta))
                    FECHAEEUU$ = Mid$(FECHA0$, 7, 4) & Mid$(FECHA0$, 4, 2) & Mid$(FECHA0$, 1, 2)
                    
                    JJ& = JJ& + 1
                    With mfgDetallePorCuenta
                        .TextMatrix(JJ&, 1) = FECHA$
                        .TextMatrix(JJ&, 2) = DOCUMENTO$
                        .TextMatrix(JJ&, 3) = REFERENCIA$
                        .TextMatrix(JJ&, 4) = IMPORTE$
                        .TextMatrix(JJ&, 5) = FECHAEEUU$
                    End With
                    
                    Acumulado# = Acumulado# + XVALO(!idebecomp) - XVALO(!ihabecomp)
                    .MoveNext
                 Loop
            End With
            On Error Resume Next
            CABTEMP.Close
            Set CABATEMP = Nothing
            On Error GoTo 0
    
            ' *** CARGA DATOS DESDE DETAASIEN (SOLO ASIENTOS MANUALES TIPOASIEN=0)
            CondicionSqlContable$ = REPLACAR$(CondicionSql$, "cuecontai", "cuecoint")
            ConsultaSql$ = "SELECT encabeasien.fechasien, encabeasien.nuidasien, encabeasien.refeasien, idebepase, ihabepase " _
            & " from detaasien with(nolock) inner join encabeasien with(nolock) on (detaasien.nuidasien=encabeasien.nuidasien) " _
            & " where tipoasien=0 and (detaasien.fechasien>= '" & FETEXCOR1$(DesdeFecha%) & "' and detaasien.fechasien <= '" & FETEXCOR1$(HastaFecha%, "HASTA") & "') " _
            & " and cuecoint=" & CStr(CuentaContable%) & " and detaasien.statpase>0 and encabeasien.statasien>0"
            Dim RstDetaasien As ADODB.Recordset
            Set RstDetaasien = READSET(ConsultaSql$)
            '
            mfgDetallePorCuenta.Rows = mfgDetallePorCuenta.Rows + RstDetaasien.RecordCount
            '
            With RstDetaasien
                Do While Not .EOF
                    
                    FECHA$ = FECHATEX$(CVINT(!FECHASIEN))
                    DOCUMENTO$ = !nuidasien
                    REFERENCIA$ = !refeasien
                    IMPORTE$ = !idebepase - !ihabepase
                        FECHA0$ = FECHATEXLAR$(CVINT(!FECHASIEN))
                    FECHAEEUU$ = Mid$(FECHA0$, 7, 4) & Mid$(FECHA0$, 4, 2) & Mid$(FECHA0$, 1, 2)
                    
                    JJ& = JJ& + 1
                    With mfgDetallePorCuenta
                        .TextMatrix(JJ&, 1) = FECHA$
                        .TextMatrix(JJ&, 2) = DOCUMENTO$
                        .TextMatrix(JJ&, 3) = REFERENCIA$
                        .TextMatrix(JJ&, 4) = IMPORTE$
                        .TextMatrix(JJ&, 5) = FECHAEEUU$
                    End With
                    
                    Acumulado# = Acumulado# + XVALO(!idebepase) - XVALO(!ihabepase)
                    
                    .MoveNext
                Loop
                
                mfgDetallePorCuenta.COL = 5
                mfgDetallePorCuenta.Sort = 1
            End With
            On Error Resume Next
            RstDetaasien.Close
            Set RstDetaasien = Nothing
            On Error GoTo 0
            
            lblTotalCuenta = TRIM$(FORMATNUM$(Acumulado#, "F12.2"))
            
        End If
        
        
End Sub

Private Sub mnuImpresionDetalleCuenta_Click()

   CommonDialog1.CancelError = True
   CommonDialog1.DialogTitle = "Exportar a Excel"
       ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
   CommonDialog1.Filter = "Excel Files(*.XLS)|*.XLS|"
   CommonDialog1.FilterIndex = 1
   On Error GoTo 99
   CommonDialog1.ShowOpen
   On Error GoTo 0
   '
   DoEvents
   Screen.MousePointer = 11
   AREXPO$ = UCase$(CommonDialog1.FileName)
   If TRIM$(AREXPO$) <> "" Then
      Call GeneraExcelDeGrilla(AREXPO$, mfgDetallePorCuenta, lblNombreInforme & " | " & lblPeriodo & " | Cuenta: " & txtCuentaContable & "  " & txtDescripcionCuenta, False)
   End If
   '
99 Exit Sub

End Sub

Private Sub mnuImpresionInforme_Click()

   CommonDialog1.CancelError = True
   CommonDialog1.DialogTitle = "Exportar a Excel"
       ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
   CommonDialog1.Filter = "Excel Files(*.XLS)|*.XLS|"
   CommonDialog1.FilterIndex = 1
   On Error GoTo 99
   CommonDialog1.ShowOpen
   On Error GoTo 0
   '
   DoEvents
   Screen.MousePointer = 11
   AREXPO$ = UCase$(CommonDialog1.FileName)
   If TRIM$(AREXPO$) <> "" Then
      Call GeneraExcelDeGrilla(AREXPO$, mfgInforme, lblNombreInforme & " | " & lblPeriodo, False)
   End If
   '
99 Exit Sub

End Sub

