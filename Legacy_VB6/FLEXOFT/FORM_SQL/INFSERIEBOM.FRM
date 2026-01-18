VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form INFSERIEBOM 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informe de Número de Serie de Bombas"
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   16740
   LinkTopic       =   "Form1"
   ScaleHeight     =   7380
   ScaleWidth      =   16740
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   15600
      Top             =   6720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2880
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CDDADA&
      Height          =   600
      Left            =   240
      TabIndex        =   2
      Top             =   6600
      Width           =   8655
      Begin VB.ComboBox CMBFILTRO 
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
         Left            =   5280
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   160
         Width           =   3135
      End
      Begin VB.TextBox TXTBUSCAR 
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
         Height          =   315
         Left            =   960
         TabIndex        =   3
         Top             =   160
         Width           =   2535
      End
      Begin VB.Label Label71 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D8BD05&
         BackStyle       =   0  'Transparent
         Caption         =   "Campo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4440
         TabIndex        =   6
         Top             =   240
         Width           =   705
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D8BD05&
         BackStyle       =   0  'Transparent
         Caption         =   "Filtrar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   585
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdSeriesBombasSearch 
      Height          =   6255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   16455
      _ExtentX        =   29025
      _ExtentY        =   11033
      _Version        =   393216
      BackColor       =   -2147483645
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Roboto Mono"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid grdSeriesBombas 
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   16455
      _ExtentX        =   29025
      _ExtentY        =   11033
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Roboto Mono"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "INFSERIEBOM.frx":0000
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   255
      Index           =   0
      Left            =   1200
      OleObjectBlob   =   "INFSERIEBOM.frx":0234
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
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
      Left            =   360
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Menu mnuTransacciones 
      Caption         =   "Transacciones"
      Begin VB.Menu mnuNuevoSerieBomba 
         Caption         =   "Nuevo Serie de Bomba"
      End
      Begin VB.Menu mnuEliminaSerieBomba 
         Caption         =   "Eliminar Serie de Bomba"
      End
   End
   Begin VB.Menu mnuConsultasyListados 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu mnuExportaSeriesBomba 
         Caption         =   "Exportar Números de Serie"
      End
      Begin VB.Menu mnuLinea1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReportesVarios 
         Caption         =   "Reportes Varios"
      End
   End
   Begin VB.Menu mnuConfiguración 
      Caption         =   "Configuración"
      Begin VB.Menu mnuAsociarReporte 
         Caption         =   "Vincular Reporte con Formulario"
      End
   End
End
Attribute VB_Name = "INFSERIEBOM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EncabezadoGrilla As String
Dim MODO As Byte ' 0 Edicion 1 Nuevo


Sub ActualizarSeriesDeBombas()

    If grdSeriesBombas.Rows > 1 Then
    
        Dim Fila As Long
        
        On Error GoTo MarcaDeBorrado
        FLEXCONN.BeginTrans
        FLEXCONN.Execute "update seriesbombas set marborra=1"
        On Error GoTo 0
        
        Dim rstSeries As ADODB.Recordset
        Set rstSeries = New ADODB.Recordset
        
        On Error GoTo GrabaNumerosDeSerie
        ConsultaSql$ = "select * from seriesbombas where marborra=1"
        rstSeries.Open FILTSQL$(ConsultaSql$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        
        With rstSeries
                For Fila = 1 To grdSeriesBombas.Rows - 1
                    .AddNew
                    
                    !Serie = grdSeriesBombas.TextMatrix(Fila, 1)
                    !FechaAsignacion = CVDAT(FECHANUM(grdSeriesBombas.TextMatrix(Fila, 2)))
                    !Codigo = CODINT(grdSeriesBombas.TextMatrix(Fila, 3))
                    !Cliente = CLng(grdSeriesBombas.TextMatrix(Fila, 5))
                    !Tag = grdSeriesBombas.TextMatrix(Fila, 7)
                    !Yacimiento = grdSeriesBombas.TextMatrix(Fila, 8)
                    !NotaDeVenta = CLng(grdSeriesBombas.TextMatrix(Fila, 9))
                    !FechaNotaDeVenta = CVDAT(FECHANUM(grdSeriesBombas.TextMatrix(Fila, 10)))
                    !ORDENCOMPRA = grdSeriesBombas.TextMatrix(Fila, 11)
                    !FechaOrdenCompra = CVDAT(FECHANUM(grdSeriesBombas.TextMatrix(Fila, 12)))
                    !CodigoCliente = grdSeriesBombas.TextMatrix(Fila, 13)
                    !Motriz = grdSeriesBombas.TextMatrix(Fila, 14)
                    !CuerpoHidraulico = grdSeriesBombas.TextMatrix(Fila, 15)
                    !Ciguenal = grdSeriesBombas.TextMatrix(Fila, 16)
                    !Corona = grdSeriesBombas.TextMatrix(Fila, 17)
                    !Pinon = grdSeriesBombas.TextMatrix(Fila, 18)
                    !Reductor = grdSeriesBombas.TextMatrix(Fila, 19)
                    !Valvula = grdSeriesBombas.TextMatrix(Fila, 20)
                    !Motor = grdSeriesBombas.TextMatrix(Fila, 21)
                    !AmortigSuccion = grdSeriesBombas.TextMatrix(Fila, 22)
                    !AmortigDescarga = grdSeriesBombas.TextMatrix(Fila, 23)
                    !CarretelSuccion = grdSeriesBombas.TextMatrix(Fila, 24)
                    !CarretelDescarga = grdSeriesBombas.TextMatrix(Fila, 25)
                    !PresostatoSuccion = grdSeriesBombas.TextMatrix(Fila, 26)
                    !PresostatoDescarga = grdSeriesBombas.TextMatrix(Fila, 27)
                    !TableroElectrico = grdSeriesBombas.TextMatrix(Fila, 28)
                    !BombaCentrifuga = grdSeriesBombas.TextMatrix(Fila, 29)
                    
                    !MARBORRA = 0
                    
                    .Update
                Next Fila
        End With
        On Error GoTo 0
               
        On Error GoTo EliminaConMarcaBorrado
        FLEXCONN.Execute "delete seriesbombas where marborra=1"
        On Error GoTo 0
        
        FLEXCONN.CommitTrans
        
        On Error Resume Next
        rstSeries.Close
        Set rstSeries = Nothing
        On Error GoTo 0
        
    End If
    
    Exit Sub
    
MarcaDeBorrado:
    Mensaje$ = Err.Description
    FLEXCONN.RollbackTrans
    On Error GoTo 0
    Call COMUNI("Problemas en Proceso Marca de Borrado.\ \" & Mensaje$)
    Exit Sub

GrabaNumerosDeSerie:
    Mensaje$ = Err.Description
    FLEXCONN.RollbackTrans
    On Error GoTo 0
    Call COMUNI("Problemas en Grabación de Series.\ \" & Mensaje$)
    Exit Sub

EliminaConMarcaBorrado:
    Mensaje$ = Err.Description
    FLEXCONN.RollbackTrans
    On Error GoTo 0
    Call COMUNI("Problemas en Proceso Eliminación de Marca.\ \" & Mensaje$)
    Exit Sub

End Sub

Private Sub Form_Load()
   
    Call ABRECONEXION
    
    EncabezadoGrilla = "Serie/A16;Asignado/D8;Código/A16;Descripción/A36;Cliente/I10;Razon Social/A36;TAG/A16;Yacimiento/A24;Nota de Vta./I15;Fecha N.Vta/D10;O-Compra/A24;Fecha O-C/D10;Cod. Cliente/A16;Motriz/A16;Cuerpo H./A16;Cigueñal/A16;Corona/A16;Piñon/A16;Reductor/A16;Válvula/A16;Motor/A16;Amort.Succión/A16;Amort.Descarga/A16;Carr.Succión/A16;Carr.Descarga/A16;Pres.Succión/A16;Pres.Descarga/A16;Tab.Eléctrico/A16;Bomba Cent./A16"
    Call CARGA_ENCABEZADO(grdSeriesBombas, EncabezadoGrilla, Me, 1)
    Call CARGA_ENCABEZADO(grdSeriesBombasSearch, EncabezadoGrilla & ";FilaGridPrincipal/A10", Me, 1)
    
    Call CargaSerieDeBombas
    
End Sub


Private Sub grdSeriesBombas_Click()

    If grdSeriesBombas.MouseRow < 1 Or grdSeriesBombas.Rows < 2 Then Exit Sub
    
    For i = 1 To grdSeriesBombas.MouseRow - 1
        grdSeriesBombas.TextMatrix(i, 0) = ""
    Next i
    For i = grdSeriesBombas.MouseRow + 1 To grdSeriesBombas.Rows - 1
        grdSeriesBombas.TextMatrix(i, 0) = ""
    Next i

    grdSeriesBombas.TextMatrix(grdSeriesBombas.MouseRow, 0) = "*"
    
End Sub

Private Sub grdSeriesBombas_DblClick()

    If (grdSeriesBombas.MouseRow < 1 And MODO = 0) Or grdSeriesBombas.Rows < 2 Then Exit Sub
    
    Unload VENTABNEW
    
    VENTANA.Inicializar = 0
    VENTANA.AgregaRegistro = 0
    VENTANA.UtilizaFormula = 1
    VENTANA.HabilitaBorrar = 0
    VENTANA.HabilitaInsertar = 0
    VENTANA.NoMostrarBotonTilde = 0
    VENTANA.TITULO = "SERIES DE BOMBAS"
      VENTANA.CampEditables = "1;2;3;5;7;8;9;10;11;12;13;14;15;16;17;18;19;20;21;22;23;24;25;26;27;28;29"
    VENTANA.ColEcoCampo(3) = "codigo;descripcion;master;status>0"
    VENTANA.ColEcoCampo(5) = "nume_cli;raso_cli;deudor12;stat_cli>0"

    VENTANA.Campox = EncabezadoGrilla
    VENTANA.Inicializar = 1
    If MODO = 0 Then RenglonElegido& = grdSeriesBombas.MouseRow ' edicion
    If MODO = 1 Then RenglonElegido& = grdSeriesBombas.Rows - 1 ' nuevo
    If MODO = 2 Then RenglonElegido& = grdSeriesBombas.Row      ' edicion con filtro
    
    Call EDITARFILA(grdSeriesBombas, VENTABNEW.MSF, RenglonElegido&)

10  VENTABNEW.Show 1
    If VENTABNEW.APROBADO% > 0 Then
        For i = 1 To 29
            Select Case i
                Case 1, 7, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29:
                        CELDA$ = Left(VENTABNEW.MSF.TextMatrix(1, i), 24)
                        
                Case 2:
                        CELDA$ = TRIM$(VENTABNEW.MSF.TextMatrix(1, i))
                        If CELDA$ <> "" Then
                            If FECHANUM(CELDA$) = 0 Then
                                Call MENSERR(24, "Fecha de Asignación Errónea.")
                                GoTo 10
                            End If
                        Else
                            CELDA$ = "01/01/80"
                        End If
                
                Case 5:
                        CELDA$ = TRIM$(VENTABNEW.MSF.TextMatrix(1, i))
                        If CELDA$ <> "" And CELDA <> "0" Then
                            If IsNumeric(CELDA$) = False Or STATCLI%(CLng(CELDA$)) = 0 Then
                                Call MENSERR(24, "Número de Cliente Inválido.")
                                GoTo 10
                            End If
                        Else
                            CELDA$ = "0"
                        End If
                        
                Case 9: CELDA$ = VENTABNEW.MSF.TextMatrix(1, i)
                        If CELDA$ <> "" And CELDA <> "0" Then
                            If IsNumeric(CELDA$) = False Then
                                Call MENSERR(24, "Nota de Venta Incorrecto.")
                                GoTo 10
                            End If
                        Else
                            CELDA$ = "0"
                        End If
                
                Case 10:
                        CELDA$ = TRIM$(VENTABNEW.MSF.TextMatrix(1, i))
                        If CELDA$ <> "" Then
                            If FECHANUM(CELDA$) = 0 Then
                                Call MENSERR(24, "Fecha de Nota de Venta Errónea.")
                                GoTo 10
                            End If
                        Else
                            CELDA$ = "01/01/80"
                        End If

                Case 12:
                        CELDA$ = TRIM$(VENTABNEW.MSF.TextMatrix(1, i))
                        If CELDA$ <> "" Then
                            If FECHANUM(CELDA$) = 0 Then
                                Call MENSERR(24, "Fecha de O-Compra Errónea.")
                                GoTo 10
                            End If
                        Else
                            CELDA$ = "01/01/80"
                        End If
                        
                Case Else:
                        CELDA$ = VENTABNEW.MSF.TextMatrix(1, i)
                        
            End Select
            
            VENTABNEW.MSF.TextMatrix(1, i) = TRIM$(CELDA$)
        Next i
        
        NUMEROSERIE$ = VENTABNEW.MSF.TextMatrix(1, 1)
        For i = 1 To grdSeriesBombas.Rows - 1
            If grdSeriesBombas.TextMatrix(i, 1) = NUMEROSERIE$ And i <> RenglonElegido& Then
                Call COMUNI("Imposible Grabar.\Número de Serie de Bomba Repetido")
                GoTo 10
            End If
        Next i
        
        For i = 1 To 29
            CELDA$ = VENTABNEW.MSF.TextMatrix(1, i)
            grdSeriesBombas.TextMatrix(RenglonElegido&, i) = CELDA$
        Next i

        Call ActualizarSeriesDeBombas
    Else
        If MODO = 1 Then
            grdSeriesBombas.Rows = grdSeriesBombas.Rows - 1
        End If
    End If
    
99  MODO = 0

End Sub

Sub CargaSerieDeBombas()

    Dim rstSeries As ADODB.Recordset
    Dim Fila As Long
    
    ConsultaSql$ = "select * from seriesbombas with(nolock) where marborra=0 order by cast(serie as int) asc"
    Set rstSeries = READSET(ConsultaSql$)
    
    Fila = 0
    With rstSeries
        If Not (.EOF And .BOF) Then
            grdSeriesBombas.Rows = .RecordCount + 1
            Do While Not .EOF
                Fila = Fila + 1
                grdSeriesBombas.TextMatrix(Fila, 1) = !Serie
                grdSeriesBombas.TextMatrix(Fila, 2) = FECHATEX$(CVINT(!FechaAsignacion))
                grdSeriesBombas.TextMatrix(Fila, 3) = CODEXT$(!Codigo)
                grdSeriesBombas.TextMatrix(Fila, 4) = DESCRIT0$(!Codigo)
                    NumeroCliente& = CLng(!Cliente)
                grdSeriesBombas.TextMatrix(Fila, 5) = CStr(NumeroCliente&)
                    A = rasocli$(0)
                grdSeriesBombas.TextMatrix(Fila, 6) = rasocli$(NumeroCliente&)
                grdSeriesBombas.TextMatrix(Fila, 7) = !Tag
                grdSeriesBombas.TextMatrix(Fila, 8) = !Yacimiento
                grdSeriesBombas.TextMatrix(Fila, 9) = !NotaDeVenta
                grdSeriesBombas.TextMatrix(Fila, 10) = FECHATEX$(CVINT(!FechaNotaDeVenta))
                grdSeriesBombas.TextMatrix(Fila, 11) = !ORDENCOMPRA
                grdSeriesBombas.TextMatrix(Fila, 12) = FECHATEX$(CVINT(!FechaOrdenCompra))
                grdSeriesBombas.TextMatrix(Fila, 13) = !CodigoCliente
                grdSeriesBombas.TextMatrix(Fila, 14) = !Motriz
                grdSeriesBombas.TextMatrix(Fila, 15) = !CuerpoHidraulico
                grdSeriesBombas.TextMatrix(Fila, 16) = !Ciguenal
                grdSeriesBombas.TextMatrix(Fila, 17) = !Corona
                grdSeriesBombas.TextMatrix(Fila, 18) = !Pinon
                grdSeriesBombas.TextMatrix(Fila, 19) = !Reductor
                grdSeriesBombas.TextMatrix(Fila, 20) = !Valvula
                grdSeriesBombas.TextMatrix(Fila, 21) = !Motor
                grdSeriesBombas.TextMatrix(Fila, 22) = !AmortigSuccion
                grdSeriesBombas.TextMatrix(Fila, 23) = !AmortigDescarga
                grdSeriesBombas.TextMatrix(Fila, 24) = !CarretelSuccion
                grdSeriesBombas.TextMatrix(Fila, 25) = !CarretelDescarga
                grdSeriesBombas.TextMatrix(Fila, 26) = !PresostatoSuccion
                grdSeriesBombas.TextMatrix(Fila, 27) = !PresostatoDescarga
                grdSeriesBombas.TextMatrix(Fila, 28) = !TableroElectrico
                grdSeriesBombas.TextMatrix(Fila, 29) = !BombaCentrifuga
        
                .MoveNext
            Loop
        End If
    End With
    
    On Error Resume Next
    rstSeries.Close
    Set rstSeries = Nothing
    On Error GoTo 0
    
End Sub

Private Sub grdSeriesBombasSearch_DblClick()
    
    If grdSeriesBombasSearch.MouseRow < 1 Then Exit Sub
    
    filaGrillaSearch& = grdSeriesBombasSearch.MouseRow
    FilaGrillaPrincipal& = grdSeriesBombasSearch.TextMatrix(filaGrillaSearch&, grdSeriesBombasSearch.Cols - 1)
    grdSeriesBombas.Visible = True
    grdSeriesBombasSearch.Visible = False
    TXTBUSCAR.TEXT = ""
    
    grdSeriesBombas.Row = FilaGrillaPrincipal&
    MODO = 2
    Call grdSeriesBombas_DblClick
    MODO = 0
    
End Sub

Private Sub mnuAsociarReporte_Click()
    Call CONECTA_CONSUL_PRE(TRIM$(UCase$(Name)))
End Sub

Private Sub mnuEliminaSerieBomba_Click()
    
    For i = 1 To grdSeriesBombas.Rows - 1
        If grdSeriesBombas.TextMatrix(i, 0) = "*" Then
            Mensaje$ = "¿Está Seguro que Desea Eliminar\La Bomba con Código '" & grdSeriesBombas.TextMatrix(i, 2) & "'\y Serie '" & grdSeriesBombas.TextMatrix(i, 12) & "'\\No, Cancelar\Si, Eliminar"
            Respuesta% = COMALTER%(Mensaje$)
            If Respuesta% = 2 Then
                grdSeriesBombas.RemoveItem (i)
                Call ActualizarSeriesDeBombas
            End If
            
            Exit For
        End If
    Next i
    
End Sub

Private Sub mnuExportaSeriesBomba_Click()

   CommonDialog1.CancelError = True
   CommonDialog1.DialogTitle = "Exportar a Excel"
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
     If grdSeriesBombas.Visible = True Then
      Call GeneraExcelDeGrilla(AREXPO$, grdSeriesBombas, "CONJUNTOS ARMADOS DE BOMBAS")
     Else
      Call GeneraExcelDeGrilla(AREXPO$, grdSeriesBombasSearch, "CONJUNTOS ARMADOS DE BOMBAS" & " - Filtro: " & TXTBUSCAR)
     End If
   End If
   '
   Exit Sub
   
99:
   Exit Sub
End Sub

Private Sub mnuNuevoSerieBomba_Click()

    MODO = 1
    grdSeriesBombas.Rows = grdSeriesBombas.Rows + 1
    Call grdSeriesBombas_DblClick

End Sub

Private Sub mnuReportesVarios_Click()

    Call ACONSUL_PRE(TRIM$(UCase$(Name)))

End Sub

Private Sub TXTBUSCAR_Change()
    '
    If TRIM$(TXTBUSCAR) = "" Then
      grdSeriesBombasSearch.Visible = False
      grdSeriesBombas.Visible = True
      Exit Sub
    End If
    '
    grdSeriesBombasSearch.Rows = 2
    For i& = 1 To grdSeriesBombasSearch.Cols - 1
       grdSeriesBombasSearch.TextMatrix(1, i&) = ""
    Next i&
    REG& = 0
    grdSeriesBombasSearch.Visible = True
    grdSeriesBombasSearch.ZOrder 0
    'grdSeriesBombasSearch.Cols = grdSeriesBombas.Cols
    grdSeriesBombasSearch.Redraw = False
    grdSeriesBombas.Visible = False
    '
    'RECORRO EL MSF QUE TIENE TODO E IMPRIMO EN msf  LOS QUE MACHEAN
    For i& = 1 To grdSeriesBombas.Rows - 1
      For J& = 1 To grdSeriesBombas.Cols - 1
        If (InStr(UCase$(grdSeriesBombas.TextMatrix(i&, J&)), UCase$(TXTBUSCAR))) Then
          If (TRIM$(CMBFILTRO.TEXT)) = "" Or TRIM$(UCase$(CMBFILTRO)) = TRIM$(UCase$(grdSeriesBombas.TextMatrix(0, J&))) Then
            REG& = REG& + 1
            grdSeriesBombasSearch.Rows = REG& + 1
            For H& = 1 To grdSeriesBombasSearch.Cols - 2
              grdSeriesBombasSearch.TextMatrix(REG&, H&) = grdSeriesBombas.TextMatrix(i&, H&)
            Next H&
            grdSeriesBombasSearch.TextMatrix(REG&, grdSeriesBombasSearch.Cols - 1) = i&
          End If
          If (TRIM$(CMBFILTRO.TEXT)) = "" Then Exit For
        End If
      Next J&
    Next i&
    '
    grdSeriesBombasSearch.Redraw = True
    '
End Sub

Private Sub TXTBUSCAR_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      TXTBUSCAR.TEXT = ""
   End If
End Sub

