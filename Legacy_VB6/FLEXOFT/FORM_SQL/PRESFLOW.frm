VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form PRESFLOW 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   8970
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   14595
   LinkTopic       =   "Form1"
   ScaleHeight     =   8970
   ScaleWidth      =   14595
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7080
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Detalle "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   30
      TabIndex        =   3
      Top             =   5280
      Width           =   14535
      Begin VB.TextBox TXTID 
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
         Height          =   315
         Left            =   400
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   210
         Width           =   975
      End
      Begin VB.TextBox TXTMODO 
         Alignment       =   2  'Center
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
         Height          =   315
         Left            =   2320
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   210
         Width           =   975
      End
      Begin VB.TextBox TxtMovimiento 
         Alignment       =   2  'Center
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
         Height          =   315
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   210
         Width           =   1575
      End
      Begin VB.TextBox TxtRango 
         Alignment       =   2  'Center
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
         Height          =   315
         Left            =   9240
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   210
         Width           =   1695
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D8BD05&
         BackStyle       =   0  'Transparent
         Caption         =   "ID:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   345
         Width           =   225
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D8BD05&
         BackStyle       =   0  'Transparent
         Caption         =   "MODO:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   1680
         TabIndex        =   10
         Top             =   345
         Width           =   585
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D8BD05&
         BackStyle       =   0  'Transparent
         Caption         =   "Movimiento:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   4920
         TabIndex        =   9
         Top             =   345
         Width           =   1185
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D8BD05&
         BackStyle       =   0  'Transparent
         Caption         =   "Rango/Fecha:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   7920
         TabIndex        =   8
         Top             =   345
         Width           =   1185
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   5205
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14520
      _ExtentX        =   25612
      _ExtentY        =   9181
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      FillStyle       =   1
      GridLines       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Roboto Mono Medium"
         Size            =   9.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSF2 
      Height          =   2900
      Left            =   0
      TabIndex        =   1
      Top             =   6000
      Width           =   14520
      _ExtentX        =   25612
      _ExtentY        =   5106
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      FillStyle       =   1
      GridLines       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Roboto Mono Medium"
         Size            =   9.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      BeginProperty Font 
         Name            =   "Roboto Mono"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mnuPrintCashMaestro 
         Caption         =   "Imprimir Casfflow Maestro"
         Begin VB.Menu mnuImpresiónGenreporteGrillaSup 
            Caption         =   "Impresión Salida x Generador de Reporte"
         End
         Begin VB.Menu mnuExportaexcelGrillasup 
            Caption         =   "Exportación Directa a Excel "
         End
      End
      Begin VB.Menu mnuPrintCas 
         Caption         =   "Imprimir Cashflow Detalle"
         Begin VB.Menu mnuPrintCashDetalle 
            Caption         =   "Impresión Salida x Generador de Reporte"
         End
         Begin VB.Menu mnuImpresiónGenreporteGrillaInf 
            Caption         =   "Exportación Directa a Excel "
         End
         Begin VB.Menu S 
            Caption         =   "-"
         End
      End
      Begin VB.Menu mnuPB 
         Caption         =   "Proyección Bancaria"
      End
   End
End
Attribute VB_Name = "PRESFLOW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Resize()
   On Error Resume Next
   Me.MSF1.Width = Me.Width - 300
   Me.MSF2.Width = Me.Width - 300
   Me.MSF2.Height = Me.Height - (MSF1.Height + Frame1.Height + 1100)
   On Error GoTo 0
End Sub

Private Sub mnuExportaexcelGrillasup_Click()
   CommonDialog1.CancelError = True
   CommonDialog1.DialogTitle = "Exportar a Excel"
       ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
   CommonDialog1.FILTER = "Excel Files(*.XLS)|*.XLS|"
   CommonDialog1.FilterIndex = 1
   On Error GoTo 99
   CommonDialog1.ShowOpen
   On Error GoTo 0
   '
   DoEvents
   Screen.MousePointer = 11
   AREXPO$ = UCase$(CommonDialog1.FileName)
   If TRIM$(AREXPO$) <> "" Then
      Call GeneraExcelDeGrilla(AREXPO$, MSF1, PRESFLOW.Caption)
   End If
   '
99
End Sub

Private Sub mnuImpresiónGenreporteGrillaInf_Click()

   CommonDialog1.CancelError = True
   CommonDialog1.DialogTitle = "Exportar a Excel"
       ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
   CommonDialog1.FILTER = "Excel Files(*.XLS)|*.XLS|"
   CommonDialog1.FilterIndex = 1
   On Error GoTo 99
   CommonDialog1.ShowOpen
   On Error GoTo 0
   '
   DoEvents
   Screen.MousePointer = 11
   AREXPO$ = UCase$(CommonDialog1.FileName)
   If TRIM$(AREXPO$) <> "" Then
      TXXX$ = "Movimiento: " & TxtMovimiento.TEXT & " - Rango/Fecha: " & TxtRango.TEXT
      Call GeneraExcelDeGrilla(AREXPO$, MSF2, TXXX$)
   End If
   '
99

End Sub


Private Sub mnuImpresiónGenreporteGrillaSup_Click()
  If MSF1.Cols < 20 Then
     Call PrintCashMaestro
  Else
     Call COMUNI("Atención !!! \Cantidad de Columnas del Reporte\Supera el Máximo Permitido\Para Salida por Generador de Reportes\Utilizar Opción: -Exportación Directa a Excel-")
  End If
End Sub

Private Sub mnuPB_Click()

'Sub AGREGARVALORESCASHFLOW(MSF As MSFlexGrid, TIPOPERIODO$, IDKABEZA)
    TIPOPERIODO$ = TXTMODO
    IDKABEZA = XVALO(TXTID)
    i& = 2:
    VENTABNEW.MSF.Rows = 1
    Dim NUCUENTBLE$(), DESCRICUENTA$(), SUBTOTAL()
    SQLX$ = "SELECT * FROM  DATASQL WITH(NOLOCK) WHERE LISTA = 'CUENTASCONBLES' "
    Set RS = READSET(SQLX$)
    With RS
      Do While Not .EOF
         i& = i& + 1
         On Error Resume Next
         If UBound(NUCUENTBLE$) < i& Then ReDim Preserve NUCUENTBLE$(i&), DESCRICUENTA$(i&), SUBTOTAL(i&)
         On Error GoTo 0
         NUCUENTBLE$(i&) = SAFETEXT$(!CAMPO1)
         DESCRICUENTA$(i&) = SAFETEXT$(!CAMPO2)
         
        .MoveNext
      Loop
    End With
    '
    COLALEER$ = "1;2"
    'IMPRIIMIR LAS COLUMNAS CON LAS CUENTAS CONTABLES DEJANDO 2 LUGARES PARA LA REFERENCIA
    Campos$ = "Referencia Movimiento/A48;/A12"
    For i& = 3 To UBound(NUCUENTBLE$)
        Campos$ = Campos$ & ";" & NUCUENTBLE$(i&) & "(" & DESCRICUENTA$(i&) & ")" & "/F32.2"
        COLALEER$ = COLALEER$ + ";" & i&
    Next i&
    Campos$ = Campos$ & ";SUTOTALES/F12.2"
    Call CARGA_ENCABEZADO(VENTABNEW.MSF, Campos$, Me, , , , , , 0)
    FIL% = 3
    
'    VER EL ULTIMO REGISTRO EN LA PROYECCION AMODO DIARIO QUE PARECE QUE HAY UNO DE MAS.
    Dim MSF As MSFlexGrid
    Set MSF = MSF1

    For i& = 1 To MSF.Cols - 1
        LEIDO$ = MSF.TextMatrix(0, i&)
        If TIPOPERIODO$ = "DIA" Then
           FEFI% = FECHANUM(LEIDO$)
           If FEFI% < 1 Then GoTo 80
           If i& = 1 Then
              Condi$ = "ACREDITACION <= '" & FETEXCOR1$(FEFI%) & "'"
           Else
              Condi$ = "ACREDITACION = '" & FETEXCOR1$(FEFI%) & "'"
           End If
        ElseIf TIPOPERIODO$ = "SEM" Then
           FEIN% = FECHANUM(Left(LEIDO$, 8))
           FEFI% = FECHANUM(Mid$(LEIDO$, 10, 8))
           If FEFI% < 1 Then GoTo 80
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
           ANO% = XVALO(Mid$(LEIDO$, 6, 4))
           If NMES% < 1 Then GoTo 80
           If i& = 1 Then
              Condi$ = " YEAR(ACREDITACION) <= " & ANO%
              Condi$ = Condi$ + " AND  MONTH(ACREDITACION) <= " & NMES%
           Else
              Condi$ = " YEAR(ACREDITACION) = " & ANO%
              Condi$ = Condi$ + " AND  MONTH(ACREDITACION) = " & NMES%
           End If
        End If
        '
        
        'Call CARGARFILA(i&, CONDI$, MSF)
         SQLX$ = "select SUM(IMPORTE) as importe, TIPODOCU,CUCONTA_BANCO,FECHVEN"
         SQLX$ = SQLX$ + " from backupflow WITH(NOLOCK)"
         SQLX$ = SQLX$ + " where id_cabezera = " & IDKABEZA & " AND TIPODOCU <> '' "
         SQLX$ = SQLX$ + " and cuconta_banco <> '' And " & Condi$
         SQLX$ = SQLX$ + " GROUP BY TIPODOCU, CUCONTA_BANCO, FECHVEN "
         SQLX$ = SQLX$ + " ORDER BY TIPODOCU"
         Set RS = READSET(SQLX$)
         With RS
             Do While Not .EOF
               DESCRIPCIONFILA$ = ""
               'LEYENDA$ = SAFETEXT$(!MOVCASHFLOW)
               LEYENDA$ = SAFETEXT$(!TIPODOCU) 'ver como queda antes movcashflow
               TIDOC$ = UCase$(TRIM$(SAFETEXT$(!TIPODOCU)))
               Select Case TIDOC$
                  Case "FVEN"
                     LEYENDA$ = "Fact.x Cobrar"
                  Case "FCOM"
                     LEYENDA$ = "Compras"
                  Case "CHENTR"
                     LEYENDA$ = "Cheq x Pagar"
                  Case "CHRECI"
                     LEYENDA$ = "Cheq x Cobrar"
                  Case "TC"
                     LEYENDA$ = "Tarj.a Acred."
                  Case "SALBANCO"
                     LEYENDA$ = "Saldo Banco"
                  Case "ME"
                     LEYENDA$ = "Medios Electrónicos"
                  End Select

               CCBLEXLEYENDA$ = SAFETEXT$(!CUCONTA_BANCO)
               IMPO# = XVALO(!IMPORTE)
               If TRIM$(CCBLEXLEYENDA$) <> "" Then
                For J& = 1 To UBound(NUCUENTBLE$)
                   If NUCUENTBLE$(J&) = CCBLEXLEYENDA$ Then
                If LEYENDA_ANT$ <> LEYENDA$ Then
                   FIL% = FIL% + 1
                End If
                        IMPO# = IMPO# * -1 'POR DEFECTO NEGATIVO
                        If LEYENDA$ = "Saldo Banco" Or LEYENDA$ = "Fact.x Cobrar" Or LEYENDA$ = "Cheq x Cobrar" Or LEYENDA$ = "Tarj.a Acred." Or LEYENDA$ = "Medios Electrónicos" Then
                            IMPO# = IMPO# * -1 'SI ENTRA ACA VUELVE A POSITIVO
                        ElseIf ESGASTOFIJO%(LEYENDA$) Then 'LOS GASTOS QUE SE INGRESAN MANUALMENTE
                            SG% = SIGNOGASTOFIJO%(LEYENDA$)
                            IMPO# = IMPO# * SG%
                            FechaAcreditacion% = CVINT(!FECHVEN)
                            OBS_LEYENDA$ = TRIM$(ObservacionTablaIngrSali$(LEYENDA$, FechaAcreditacion%))
                            If OBS_LEYENDA$ <> "" Then
                               LEYENDA$ = LEYENDA$ & " (" & OBS_LEYENDA$ & ")"
                            End If
                        End If
                        If VENTABNEW.MSF.Rows <= FIL% Then VENTABNEW.MSF.Rows = FIL% + 1
                        If VENTABNEW.MSF.Cols <= J& Then VENTABNEW.MSF.Cols = J& + 1
                        VENTABNEW.MSF.TextMatrix(FIL%, 1) = LEYENDA$
                        VENTABNEW.MSF.TextMatrix(FIL%, J&) = TRIM$(FORMATNUM$(IMPO#, "F15.2"))
                        If IMPO# < 0 Then JJJ% = J&: Call COLOREAR_TEXTO_SOLO_UNA_AFILA(VENTABNEW.MSF, vbRed, FIL%, JJJ%)
                        SUBTOTAL(J&) = SUBTOTAL(J&) + IMPO#
                        Exit For
                   End If
                Next J&
'                FIL% = FIL% + 1
               End If
               LEYENDA_ANT$ = LEYENDA$
               .MoveNext
              Loop
           End With
           FIL% = FIL% + 1
           If VENTABNEW.MSF.Rows <= FIL% Then VENTABNEW.MSF.Rows = FIL% + 1
           VENTABNEW.MSF.TextMatrix(FIL%, 1) = "Saldo Proyectado al " & FECHATEX$(FEFI%)
           
           For ST& = 3 To UBound(NUCUENTBLE$)
             VENTABNEW.MSF.TextMatrix(FIL%, ST&) = FORMATNUM$(XVALO(SUBTOTAL(ST&)), "F12.2")
             If SUBTOTAL(ST&) < 0 Then STT% = ST&: Call COLOREAR_TEXTO_SOLO_UNA_AFILA(VENTABNEW.MSF, vbRed, FIL%, STT%)
             ACUSUBTOTAL = ACUSUBTOTAL + XVALO(SUBTOTAL(ST&))
           Next ST&
           VENTABNEW.MSF.TextMatrix(FIL%, VENTABNEW.MSF.Cols - 1) = FORMATNUM$(ACUSUBTOTAL, "F12.2")
           If ACUSUBTOTAL < 0 Then Call COLOREAR_TEXTO_SOLO_UNA_AFILA(VENTABNEW.MSF, vbRed, FIL%, VENTABNEW.MSF.Cols - 1)
           
           Call COLOREAR_GRILLA_SOLO_UNA_AFILA(VENTABNEW.MSF, &H808080, FIL%)
           FIL% = FIL% + 1
           If VENTABNEW.MSF.Rows <= FIL% Then VENTABNEW.MSF.Rows = FIL% + 1
           Call COLOREAR_GRILLA_SOLO_UNA_AFILA(VENTABNEW.MSF, &HE0E0E0, FIL%)
           'LIMPIO LOS ACUMULADORES
'           For ST& = 3 To UBound(NUCUENTBLE$)
'             SUBTOTAL(ST&) = 0
'           Next ST&
           ACUSUBTOTAL = 0
80  Next i&
    VENTABNEW.WindowState = 2
    VENTABNEW.Caption = "Proyección Bancaria"
    VENTABNEW.Show 1
    If VENTABNEW.APROBADO > 0 Then
      OPX% = COMALTER%("Opciones de Trabajo\\Cancelar\Generar Archivo Excel")
      If OPX% = 2 Then
        NombreDeArchivo$ = NOMBREDEARCHIVOEXCEL$
        If NombreDeArchivo$ <> "" Then
         Call GeneraExcelDeGrilla(NombreDeArchivo$, VENTABNEW.MSF, "Informe CashFlow Modo Bancario")
        End If
      End If
    End If
    End Sub
   Function NOMBREDEARCHIVOEXCEL$()
   CommonDialog1.CancelError = True
   CommonDialog1.DialogTitle = "Exportar a Excel"
       ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
   CommonDialog1.FILTER = "Excel Files(*.XLS)|*.XLS|"
   CommonDialog1.FilterIndex = 1
   On Error GoTo 99
   CommonDialog1.ShowOpen
   On Error GoTo 0
   '
   DoEvents
   Screen.MousePointer = 11
   NOMBREDEARCHIVOEXCEL$ = UCase$(CommonDialog1.FileName)
   '
99 End Function




Function SIGNOGASTOFIJO%(LEYENDA$)
   'DEVUELVE EL SIGNO CALCULANDO SI TIENE IMPORTE DEL LADO DEL INGRESO O DE LA SALIDA DE LOS GASTOS FIJOS
   SQLX$ = "select CAMPO3,CAMPO4 from datasql WITH(NOLOCK) where lista = 'INGRSALI' AND CAMPO3 = '" & LEYENDA$ & "'"
   Set RS = READSET(SQLX$)
   INGRESO = XVALO(RS!CAMPO3)
   SALIDA = XVALO(RS!CAMPO4)
   Resultado = INGRESO - SALIDA
   If Resultado < 0 Or LEYENDA = "Compras" Then
      SIGNOGASTOFIJO% = -1
   Else
      SIGNOGASTOFIJO% = 1
   End If
   
   
End Function
Function ObservacionTablaIngrSali$(LEYENDA$, FechaAcreditacion%)
   'DEVUELVE EL SIGNO CALCULANDO SI TIENE IMPORTE DEL LADO DEL INGRESO O DE LA SALIDA DE LOS GASTOS FIJOS
   SQLX$ = "select CAMPO6 from datasql WITH(NOLOCK) where lista = 'INGRSALI' AND CAMPO3 = '" & LEYENDA$ & "' AND CAMPOFE1='" & FETEXCOR1$(FechaAcreditacion%) & "'"
   Set RS = READSET(SQLX$)
   ObservacionTablaIngrSali$ = SAFETEXT$(RS!CAMPO6)
   RS.Close: Set RS = Nothing
   
End Function




Function ESGASTOFIJO%(LEYENDA$)
   SQLX$ = "lista = 'INGRSALI' AND CAMPO3 = '" & LEYENDA$ & "'"
   ESGASTOFIJO% = XVALO(CantRegistros("DATASQL", Condi$, "NOMESS"))
   
   
End Function

Private Sub mnuPrintCashDetalle_Click()
        If MSF2.Cols > 19 Then
           Call COMUNI("Atención !!! \Cantidad de Columnas del Reporte Supera el Máximo Permitido\Para Salida por Generador de Reportes\\Utilizar Opción: -Exportación Directa a Excel-")
           Exit Sub
        End If

        Campox$ = ""
        Select Case TRIM$(UCase(Me.TXTMODO))
          Case "DIA"
           FORMAX$ = "F12.2"
          Case Else
           FORMAX$ = "F18.2"
        End Select
        For i& = 1 To MSF2.Cols - 1
          If i& = 1 Then
           Campox$ = Campox$ & MSF2.TextMatrix(0, i&) & "/" & "A20"
          ElseIf i& = 2 Or i& = 3 Or i& = 8 Then
           If TRIM$(MSF2.TextMatrix(0, i&)) <> "" Then
              Campox$ = Campox$ & ";" & MSF2.TextMatrix(0, i&) & "/" & "A24"
           End If
          ElseIf i& = 4 Or i& = 5 Then
           If TRIM$(MSF2.TextMatrix(0, i&)) <> "" Then
              Campox$ = Campox$ & ";" & MSF2.TextMatrix(0, i&) & "/" & "F10"
           End If
          ElseIf i& = 6 Then
           If TRIM$(MSF2.TextMatrix(0, i&)) <> "" Then
              Campox$ = Campox$ & ";" & MSF2.TextMatrix(0, i&) & "/" & FORMAX$
           End If
          ElseIf i& = 7 Then
             Campox$ = Campox$ & ";" & MSF2.TextMatrix(0, i&) & "/" & "A30"
          End If
        Next i&
        Call CARGA_TABLA_IMPRE(MSF2, Campox$, "CASHMAES", Me.Caption, "TABLA_IMPRE", RET_NAMECONS$, 1, 0, 0)
        Call FINAL("?" & RET_NAMECONS$)
        
End Sub

 Sub PrintCashMaestro()
        If MSF1.Rows > 1 Then Set MSF = MSF1
        Select Case TRIM$(UCase(Me.TXTMODO))
          Case "DIA"
           FORMAX$ = "F12.2"
          Case Else
           FORMAX$ = "F18.2"
        End Select
        Campox$ = ""
        For i& = 0 To MSF1.Cols - 1
          If i& = 0 Then
           Campox$ = Campox$ & "TIPO MOVIMIENTO" & "/" & "A24"
          Else
           If TRIM$(MSF1.TextMatrix(0, i&)) <> "" Then
             Campox$ = Campox$ & ";" & MSF1.TextMatrix(0, i&) & "/" & FORMAX$
           End If
          End If
        Next i&
        Call CARGA_TABLA_IMPRE(MSF1, Campox$, "CASHMAES", Me.Caption, "TABLA_IMPRE", RET_NAMECONS$, 1, 0, -1)
        Call FINAL("?" & RET_NAMECONS$)
End Sub

Private Sub MSF1_DblClick()
   i& = MSF1.MouseRow
   J& = MSF1.MouseCol
   '
   If i& < 1 Then Exit Sub
   '
   textencabeza$ = MSF1.TextMatrix(0, J&)
   KONDIRANGO$ = CONDIRANGO$(TXTMODO, textencabeza$, J&)
   
   titulox$ = TRIM$(UCase(MSF1.TextMatrix(i&, 0)))
   Me.TxtMovimiento = titulox$
   Me.TxtRango = textencabeza$
   Select Case titulox$
      Case "CHEQ X COBRAR"
         ARCUBUS$ = "CHRECI"
      Case "FACT.X COBRAR"
         ARCUBUS$ = "FVEN"
      Case "COMPRAS"
         ARCUBUS$ = "FCOM"
      Case "CHEQ X PAGAR"
         ARCUBUS$ = "CHENTR"
      Case "TARJ.A ACRED."
          ARCUBUS$ = "TC"
      Case "Medios Electrónicos"
          ARCUBUS$ = "ME"
      Case "SALDO BANCO"
          ARCUBUS$ = "SALBANCO"
      Case Else 'EN EL CASO DE LAS CUENTAS DE IMPUESTOS O DE GASTOS FIJOS TOMA EL MISMO NOMBRE PARA EL TIPODOCU
          ARCUBUS$ = titulox$
   End Select
   '
   Call CARGA_ENCABEZADO(MSF2, "Documento/a20;F.Emisión./d12;F.Venc./d12;Cta/f7;R.Social/A24;Importe/f12.2;Desc.Mov./A24;F.Acr./D12;NC O NP/A0", Me)
   REG& = 0
   SQLX$ = "SELECT * FROM BACKUPFLOW WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE TIPODOCU = '" & ARCUBUS$ & "' AND ID_CABEZERA = " & XVALO(TXTID)
   SQLX$ = SQLX$ + " AND " & KONDIRANGO$
'   FALTA AGREGAR EN LA CONDICION QUE VERIFIQUE SI ES MES ,SEMANA O DIA Y AGREGAR ESTE DATO COMO FILTRO.
'   VER CODIGO DE AGREGARVALORESCASHFLOW Y TENER EN CUENTA QUE SE AGREGO UNA CAJA DE TEXTO TXTMODO
   Set RS = READSET(SQLX$)
   With RS
     Do While Not .EOF
        REG& = REG& + 1
        MSF2.Rows = REG& + 1
        MSF2.TextMatrix(REG&, 1) = SAFETEXT$(!DOCUORIGEN)
        MSF2.TextMatrix(REG&, 2) = FECHATEX$(CVINT(!FECHEMISI))
        MSF2.TextMatrix(REG&, 3) = FECHATEX$(CVINT(!FECHVEN))
        NP = XVALO(!NUPROVE)
        NC = XVALO(!NUCLIEN)
        If NP > 0 Then
          MSF2.TextMatrix(REG&, 4) = XVALO(NP)
          MSF2.TextMatrix(REG&, 5) = rasocli$(-1 * NP)
        Else
          MSF2.TextMatrix(REG&, 4) = XVALO(NC)
          MSF2.TextMatrix(REG&, 5) = rasocli$(NC)
        End If
        MSF2.TextMatrix(REG&, 6) = FORMATNUM$(XVALO(!IMPORTE), "F14.2")
        MSF2.TextMatrix(REG&, 7) = SAFETEXT$(!DESCRIMOV)
        MSF2.TextMatrix(REG&, 8) = FECHATEX$(CVINT(!ACREDITACION))
        MSF2.TextMatrix(REG&, 9) = NC ' guardo el ncliente si nc es 0 entonces es proveedor
        .MoveNext
     Loop
   End With
   
   
End Sub

Function CONDIRANGO$(TIPOPERIODO$, TIMECABEZA$, NUMCOLUM&)
        If TIPOPERIODO$ = "DIA" Then
           FE% = FECHANUM(TIMECABEZA$)
           If NUMCOLUM& = 1 Then
              CONDIRANGO$ = " ACREDITACION <= '" & FETEXCOR1$(FE%) & "'"
           Else
              CONDIRANGO$ = " ACREDITACION = '" & FETEXCOR1$(FE%) & "'"
           End If
           
        ElseIf TIPOPERIODO$ = "SEM" Then
           FEIN% = FECHANUM(Left(TIMECABEZA$, 8))
           FEFI% = FECHANUM(Mid$(TIMECABEZA$, 10, 8))
           If NUMCOLUM& = 1 Then
              CONDIRANGO$ = " ACREDITACION <= '" & FETEXCOR1$(FEFI%) & "'"
           Else
              CONDIRANGO$ = " ACREDITACION >= '" & FETEXCOR1$(FEIN%) & "' AND ACREDITACION <= '" & FETEXCOR1$(FEFI%) & "' "
           End If
        ElseIf TIPOPERIODO$ = "MES" Then
           NMES% = NUMMES_TRESCARACTERES%(Left(TIMECABEZA$, 3))
           ANO% = XVALO(Mid$(TIMECABEZA$, 6, 4))
           DIAX% = 0
           If Len(TIMECABEZA$) >= 12 Then DIAX% = XVALO(Mid$(TIMECABEZA$, 11, 2))
           If NUMCOLUM& = 1 Then
              CONDIRANGO$ = " YEAR(ACREDITACION) <= " & ANO%
              CONDIRANGO$ = CONDIRANGO$ + " AND  MONTH(ACREDITACION) <= " & NMES%
              If DIAX% > 0 Then CONDIRANGO$ = CONDIRANGO$ + " AND DAY(ACREDITACION) >= " & DIAX%
           Else
              CONDIRANGO$ = " YEAR(ACREDITACION) = " & ANO%
              CONDIRANGO$ = CONDIRANGO$ + " AND  MONTH(ACREDITACION) = " & NMES%
              If DIAX% > 0 Then CONDIRANGO$ = CONDIRANGO$ + " AND DAY(ACREDITACION) <= " & DIAX%
           End If
        End If
End Function

Private Sub MSF2_DblClick()
   i& = MSF2.MouseRow
   '
   If i& < 1 Then Exit Sub
   '
   DOCU$ = MSF2.TextMatrix(i&, 1)
   NC = MSF2.TextMatrix(i&, 4)
   FE% = FECHANUM(MSF2.TextMatrix(i&, 2))
   Select Case UCase$(Left$(DOCU$, 3))
     Case "RB-"
        TIDO$ = "RECIBO"
        NRO = XVALO(Mid$(DOCU$, 11, 8))
        Condi$ = "NUMECOMP= " & NRO & " AND fechemisi = '" & FETEXCOR1$(FE%) & "' and TIPOMOVIM = 'RCOB'"
        FE% = CVINT(VALOBADA("CAJABANC", "FeHoReal", Condi$, "NOMESS"))
        VARCONTROL$ = FECHATEX$(FE%)
        Call MUESTRADOC(DOCUNU&, "", TIDO$, SAFETEXT$(NRO), SAFETEXT$(NRO), VARCONTROL$)
     Case "FC-", "NC-", "ND-"
        If XVALO(MSF2.TextMatrix(i&, 9)) > 0 Then 'factura cliente
           Call MUESTRASTSCREEN(XVALO(NC), DOCU$)
        Else 'factura proveedor
           Call COMUNI("Factura de Proveedor, No posee Imagen Digital")
        End If
     Case "OP-"
        TIDO$ = "Orden de Pago"
        NRO = XVALO(Mid$(DOCU$, 11, 8))
        DOCUNU& = DOCID&(TIDO$, SAFETEXT$(NRO), SAFETEXT$(NRO))
        Call MUESTRADOC(DOCUNU&)
   End Select
   
End Sub


