VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form TRIGCONTROL 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control de Triggers y Stored Procedure"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   16185
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   16185
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "REPARAPDOCLIST"
      Height          =   375
      Left            =   5520
      TabIndex        =   20
      Top             =   120
      Width           =   1815
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   13920
      ScaleHeight     =   75
      ScaleWidth      =   2145
      TabIndex        =   17
      Top             =   360
      Width           =   2205
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   -120
         TabIndex        =   18
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.TextBox TXTSTORED 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   7440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Top             =   600
      Width           =   8700
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   600
      Left            =   80
      TabIndex        =   8
      Top             =   7200
      Width           =   7455
      Begin VB.TextBox TXTBUSCAR 
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
         Left            =   720
         TabIndex        =   10
         Top             =   160
         Width           =   2535
      End
      Begin VB.ComboBox CMBFILTRO 
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
         Left            =   4200
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   160
         Width           =   3135
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D8BD05&
         BackStyle       =   0  'Transparent
         Caption         =   "Filtrar:"
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
         TabIndex        =   12
         Top             =   160
         Width           =   585
      End
      Begin VB.Label Label71 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D8BD05&
         BackStyle       =   0  'Transparent
         Caption         =   "Campo:"
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
         Left            =   3360
         TabIndex        =   11
         Top             =   160
         Width           =   705
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exportar a Excel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   14520
      TabIndex        =   7
      Top             =   7320
      Width           =   1600
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7680
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture999 
      Height          =   225
      Index           =   0
      Left            =   1440
      ScaleHeight     =   165
      ScaleWidth      =   1110
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "TRIGCONTROL.frx":0000
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   120
      OleObjectBlob   =   "TRIGCONTROL.frx":0234
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1485
   End
   Begin MSFlexGridLib.MSFlexGrid GRID 
      Height          =   3645
      Left            =   50
      TabIndex        =   2
      ToolTipText     =   "Un Click Sobre Encabezado Para Ordenar/Doble Click Para Pasar a Grilla Inferior/Click Boton Derecho Imagen Digital"
      Top             =   3600
      Width           =   16035
      _ExtentX        =   28284
      _ExtentY        =   6429
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSF_2 
      Height          =   3645
      Left            =   50
      TabIndex        =   4
      ToolTipText     =   "Un Click Sobre Encabezado Para Ordenar/Doble Click Para Pasar a Grilla Inferior/Click Boton Derecho Imagen Digital"
      Top             =   3600
      Visible         =   0   'False
      Width           =   16035
      _ExtentX        =   28284
      _ExtentY        =   6429
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   8421631
      BackColorBkg    =   12632256
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
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
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   2565
      Left            =   120
      TabIndex        =   13
      ToolTipText     =   "Un Click Sobre Encabezado Para Ordenar/Doble Click Para Pasar a Grilla Inferior/Click Boton Derecho Imagen Digital"
      Top             =   600
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   4524
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
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   3240
      Width           =   585
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Stored Procedure/Triggers "
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Código Stored Procedure/Triggers "
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7440
      TabIndex        =   15
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Height          =   225
      Left            =   8520
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Label5"
      Height          =   240
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
   End
   Begin VB.Menu mnutrigger 
      Caption         =   "Triggers"
      Begin VB.Menu mnuvertrigger 
         Caption         =   "Visualizar Triggers"
      End
   End
   Begin VB.Menu mnusp 
      Caption         =   "Stored Procedure"
      Begin VB.Menu mnuVisualizarSP 
         Caption         =   "Visualizar Store Procedure"
      End
   End
   Begin VB.Menu mnuCont 
      Caption         =   "Control de Procedimientos"
      Begin VB.Menu mnuactualizaTablaSalotedep 
         Caption         =   "Actualizar Tabla Salotedep desde Movisto"
      End
      Begin VB.Menu mnuasignaloteamovimRTsinlote 
         Caption         =   "Asignar Lotes a Salidas RT Sin Lote"
      End
      Begin VB.Menu mnuAsignaLoteaMovimNotaCredito 
         Caption         =   "Asigna Lotes Totales a Ingreso x N.Crédito"
      End
      Begin VB.Menu mnuAsignaLotesPArcialesoTotalesaNC 
         Caption         =   "Asigna Lotes Totales o Parciales a Notas de Credito"
      End
      Begin VB.Menu mnuControlProcActivo 
         Caption         =   "Control de Proceso Activo"
      End
      Begin VB.Menu agregaloteaajuste 
         Caption         =   "AgregarLote a Ajustes"
      End
   End
   Begin VB.Menu mnuControl 
      Caption         =   "Controles"
      Begin VB.Menu mnumovistosalotedep 
         Caption         =   "Ver Todos Los Códigos Movisto VS Salotedep"
      End
      Begin VB.Menu mnuCopiaSalotedepASalotedep1 
         Caption         =   "Copiar Tabla Salotedep a Salotedep1 Para Control"
      End
      Begin VB.Menu mnuControlSalotedep1 
         Caption         =   "Ver Todos los Códigos Movisto VW Salotedep1"
      End
   End
End
Attribute VB_Name = "TRIGCONTROL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub BOTEXIT_Click()
  Unload Me
End Sub

Private Sub agregaloteaajuste_Click()
    Call AgregaLoteaAjusteSinLote_AHP
End Sub
Private Sub CMBFILTRO_Change()
   TXTBUSCAR_Change
End Sub

Private Sub Command1_Click()
   command1.Enabled = False
   '
   
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
     If MSF_2.Visible = True Then
      Call GENEXCEL(AREXPO$, MUESTRADATOS.MSF_2)
     Else
      Call GENEXCEL(AREXPO$, MUESTRADATOS.GRID)
     End If
   End If
   '

99 command1.Enabled = True

End Sub
Sub GENEXCEL(NombreDeArchivo$, GRID As msFlexGrid)
   Dim XLAPP As Excel.Application
   Dim XLBOOK As Excel.Workbook
   Dim XLSHEET As Excel.Worksheet
   '
   Dim FECHA As Date
   Dim valor As Double
   Dim Texto As String
   '
   On Error GoTo 20
   Set XLAPP = New Excel.Application
   Set XLBOOK = XLAPP.Workbooks.Add
   Set XLSHEET = XLBOOK.Worksheets.Add
   On Error GoTo 0
   '
   TipoDato% = 0: COLUM& = 0
   For i& = 1 To GRID.Rows
       RENG1& = i&: COLUM& = 0
       For COLU1& = 1 To GRID.Cols
          If GRID.ColWidth(COLU1& - 1) = 0 Then GoTo 10 ' OBVIA LAS COLUMNAS OCULTAS
          COLUM& = COLUM& + 1
          '
          VALORX = GRID.TextMatrix(i& - 1, COLU1 - 1)
          If Len(VALORX) = 8 And Mid$(VALORX, 3, 1) = "/" And Mid$(VALORX, 6, 1) = "/" Then
             FEXU$ = VALORX
             VALORX = FETEXCOR1$(Int(FECHANUM(FEXU$)))
             TipoDato% = 1
          End If
          ' ESTO ES PARA QUE EXPORTE COMO CERO Y NO ESPACIOS
          FFCC$ = UCase$(Left$(FORCAMPO$(COLU1&), 1))
          If FFCC$ = "I" Or FFCC$ = "F" Or FFCC$ = "G" Or FFCC$ = "$" Then
             If TRIM$(CStr(VALORX)) = "" Then
                VALORX = 0
             End If
             TipoDato% = 2
          End If
          If FFCC$ = "C" Or FFCC$ = "A" Then VALORX = "'" & VALORX
          '
          'PARA QUE NO REPITA ENCABEZADO
'          If RENG1& = 1 Then
'            If VALORX = VALORX_ANT And TRIM$(VALORX) <> "" Then
'              If COLUM& <= 17 Then
'               LETRAX1$ = Chr$(64 + COLUM& - 1)
'               LETRAX2$ = Chr$(64 + COLUM&)
'               RANXO$ = LETRAX1$ & TRIM$(Str$(RENG1&)) & ":" & LETRAX2$ & TRIM$(Str$(RENG1&))
'               Range(RANXO$).Select
'
'               XLSHEET.Cells.HorizontalAlignment = xlCenter
'               XLSHEET.Cells.HorizontalAlignment = xlCenter
'               XLSHEET.Cells.VerticalAlignment = xlCenter
'               XLSHEET.Cells.WrapText = False
'               XLSHEET.Cells.Orientation = 0
'               XLSHEET.Cells.AddIndent = False
'               XLSHEET.Cells.IndentLevel = 0
'               XLSHEET.Cells.ShrinkToFit = False
'               XLSHEET.Cells.ReadingOrder = xlContext
'               XLSHEET.Cells.MergeCells = True
'               GoTo 10
'             End If
'            End If
'          End If
          '
          XLSHEET.Cells(RENG1&, COLUM&) = VALORX
          If TipoDato% = 1 Then XLSHEET.Cells(RENG1&, COLUM&).NumberFormat = "dd/mm/yyyy"
          'If TIPODATO% = 2 Then XLSHEET.Cells(RENG1&, COLU1&).NumberFormat = "#,##0.00"
          If TipoDato% = 0 Then XLSHEET.Cells(RENG1&, COLUM&).NumberFormat = "@"
          TipoDato% = 0
          VALORX_ANT = VALORX
10     Next COLU1&
   Next i&
   '
   Close #N1%
   On Error GoTo 30
   XLSHEET.SaveAs NombreDeArchivo$
   On Error GoTo 0
   '
   XLAPP.Quit      ' cierra EXCEL que se esta ejecutando minimizado
   Set XLAPP = Nothing     ') libera memoria
   Set XLBOOK = Nothing    ') utilizada por
   Set XLSHEET = Nothing   ') las variables objeto
   '
   Screen.MousePointer = 1
   Call COMUNI("Planilla Excel Creada Exitosamente:\ \" + REPLACAR$(NombreDeArchivo$, "\", "/"))
   '
Exit Sub
   '
   '
20 Resume 21
21 On Error GoTo 0
   MsgBox "Utilitario EXCEL no Instalado"
   On Error GoTo 0
   Exit Sub
   '
30 Resume 31
31 On Error GoTo 0
   Screen.MousePointer = 1
   '
   ' Cierra el objeto Book, el cual es utilizado al crear
   ' una nueva hoja en el archivo Excel.
   ' El valor False es para que no pregunte si quiere
   ' guardar el archivo.
   XLBOOK.Close (False)
   XLAPP.Quit ' Cierra Excel
   Set XLAPP = Nothing
   Set XLBOOK = Nothing
   Set XLSHEET = Nothing
   '
   Call MENSERR(24, "Imposible Crear Planilla Excel.")
   Exit Sub
   '
End Sub



Private Sub Command3_Click()

End Sub






Private Sub Command2_Click()
  Command2.Enabled = False
  'ESTA RUTINA IGUALA EL NUMERO DE POSICION DEL REGISTRO DE ARCHIVO PDOCLIST CON EL CAMPO NUMERICO DE BUSQUEDA
  'SE USO PARA AHP POR QUE SE HABIA DEFASADO LA NUMERACION CON LA POSICION (HABIA EMPEZADO A NUMERAR DE 1)
  If CLACONTRA < 1 Then GoTo 99
20 Call BLOQARCH("PDOCLST")
   If Err Then
21    On Error GoTo 0
      Call COMUNI("Error al Bloquear Tablas'PDOCLST'.\  \Espere un Instante y Pulse 'Aceptar'\para Reanudar Ejecución.")
      GoTo 20
   End If
   On Error GoTo 0
   Call ABRECONEXION
   UREGI& = ULTREG&("PDOCLST")
   For i& = 27529 To UREGI&   '27528 ES A APARTIR DE QUE NUMERACION EMPEZO A NUMERAR DE 1 ASI QUE EL PROXIMO 27529
     TX$ = REGLEIDO$("PDOCLST", i&)
     DOCU& = CVS(Left$(REGLEIDO$("PDOCLST", i&), 4))
     Call REPLA(TX$, MKS$(i&), 1, 4)
     Call GRAREG("PDOCLST", TX$, i&)
     
     CONDI$ = "NUMEDOC = " & (i& - 27528)
     CONDI$ = CONDI$ + " And FechayHora > '01/ABR/2017' "
     Call ACTUREGISTRO("PDOC_TBL", "NUMEDOC", CONDI$, i&, REGAF&, "NOMESS")
   Next i&

   Call CIERRARCH("*.*")
   Call LIBARCH("PDOCLST")
99 Command2.Enabled = True

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If Cancel = 0 Then
     Call FINAL("")
   End If
End Sub

Private Sub mnuactualizaTablaSalotedep_Click()
    Call ABRESTOCKS
    
    If CLACONTRA% > 0 Then
        Call COMUNI("Atención !!! Si Acepta Tendrá La Ventana de Selección de Código\Si Selecciona Código, Se Actualizará Solo Los Saldos de Dicho Código\Para Correr el Proceso de Actualización de Todos Los Códigos\No Seleccione y Presione Aceptar")
10      OBJ3$ = OBJPLA$("PRESCODIGO", VDEF$)
        
        If OBJ3$ <> "" Then
            CI1% = CVI(Mid$(OBJ3$, 1, 2))
            'VALIDACIONES
            If CI1% <= 0 Then
              OPX% = COMALTER%("No se Ha ingresado Código o El Mismo es Inexistente\ Si Selecciona Correr Proceso \Actualizará Desde El Movisto La Tabla Salotedep\\Cancelar\Correr Proceso de Actualización Para Todos Los Códigos")
              If OPX% < 2 Then GoTo 10: Exit Sub
               
            End If
       Else
            Exit Sub
       End If
       Screen.MousePointer = 11
       Call PoneEnCeroNull("MOVISTO", "CANTINGRE", REGAF&, "NOMESS")
       Call PoneEnCeroNull("MOVISTO", "CANTSALID", REGAF&, "NOMESS")
       On Error GoTo ERROR_GUARDAR
       If CI1% > 0 Then
         SQLX$ = "DELETE FROM SALOTEDEP WHERE COD_INTE = " & CI1%
       Else
         SQLX$ = "DELETE FROM SALOTEDEP"
       End If

       FLEXCONN.Execute SQLX$
       SQLX$ = "SELECT SUM(CANTINGRE) AS INGRE,SUM(CANTSALID)AS SALI,COD_INTE,IDENLOTE,DEPOMOVI,CODIEMPR FROM MOVISTO WITH(NOLOCK)"
       SQLX$ = SQLX$ + " WHERE COD_INTE > 0  and CODIEMPR =  " & CodiEmp%
       If CI1% > 0 Then SQLX$ = SQLX$ + " AND COD_INTE = " & CI1%
       SQLX$ = SQLX$ + " GROUP BY COD_INTE,IDENLOTE,DEPOMOVI,CODIEMPR"
       SQLX$ = SQLX$ + " ORDER BY COD_INTE,IDENLOTE"
       Set rs = READSET(SQLX$)
       With rs
        Do While Not .EOF
            CI2& = XVALO(!cod_inte)
            CIX1% = CI2&
            SALDO = TRIM$(FORMATNUM$(XVALO(!INGRE) - XVALO(!sali), "F16." & CNTDC$(CIX1%)))
            If SALDO = "" Then SALDO = 0
            LOTE$ = SAFETEXT$(!idenlote)
            DEPO% = XVALO(!depomovi)
            FECHALTAX = FETEXCOR1$(FECREALOT%(LOTE$))
            
            NPROV = PROVELOTE%(LOTE$)
            'AGREGA EL REGISTRO
            SQLX2$ = "INSERT INTO SALOTEDEP"
            SQLX2$ = SQLX2$ + " (COD_INTE,IdenLote,DepoMovi,CODIEMPR,SALDO,FechMov,Nume_Cli)"
            SQLX2$ = SQLX2$ + " VALUES(" & CI2&
            SQLX2$ = SQLX2$ + " ,'" & LOTE$ & "'"
            SQLX2$ = SQLX2$ + " ," & DEPO%
            SQLX2$ = SQLX2$ + " ," & CodiEmp%
            SQLX2$ = SQLX2$ + " ," & SALDO
            SQLX2$ = SQLX2$ + " ,'" & FECHALTAX & "'"
            SQLX2$ = SQLX2$ + " ," & NPROV
            SQLX2$ = SQLX2$ + ")"
            FLEXCONN.Execute (SQLX2$)
          .MoveNext
        Loop
       End With
    
    
ERROR_GUARDAR:
    If Err <> 0 Then
       Screen.MousePointer = 1
       On Error Resume Next
       On Error GoTo 0
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION " + Err.Description)
       Exit Sub
    Else
      Screen.MousePointer = 1
      Call COMUNI("Proceso Realizado Exitosamente")
      rs.Close
      Set rs = Nothing
    End If
 End If
    
'   Call CREACAMPO("STOCKS", "SALOTEDEP", "COD_INTE", 4, 0)
'   Call CREACAMPO("STOCKS", "SALOTEDEP", "CODIEMPR", 3, 0)
'   Call CREACAMPO("STOCKS", "SALOTEDEP", "IdenLote", 10, 16, 1)
'   Call CREACAMPO("STOCKS", "SALOTEDEP", "SALDO", 7, 0)
'   Call CREACAMPO("STOCKS", "SALOTEDEP", "DepoMovi", 3, 0)


End Sub

Private Sub mnuAsignaLoteaMovimNotaCredito_Click()
  Call RecuperarLoteNotaDeCredito
  Call COMUNI("FINALIZO EL PROCESO")
End Sub

Private Sub mnuasignaloteamovimRTsinlote_Click()
    
    Call RecuperarLote
    Call COMUNI("FINALIZO EL PROCESO")
End Sub

Private Sub mnuAsignaLotesPArcialesoTotalesaNC_Click()
    Call RecuperarLoteNotaDeCreditoParciales
    Call COMUNI("FINALIZO EL PROCESO")
End Sub

Private Sub mnuControlProcActivo_Click()
   For i& = 1 To MSF.Rows - 1
     If MSF.TextMatrix(i&, 0) = "*" Then REG& = i&
   Next i&
   NOMB$ = TRIM$(UCase$(MSF.TextMatrix(REG&, 1)))
   If NOMB$ = "DIS_SALDOLOTE_DELETE" Or NOMB$ = "DIS_SALDOLOTE_UPDATE" Or NOMB$ = "DIS_SALDOLOTE_INSERTAR" Then
     OPX% = COMALTER%("Seleccionar Controles\\Control de Saldos X Items\Control de Saldos X Items-Depósitos\Control de Saldos X Items-Lote-Deposito")
     If OPX% < 1 Then Exit Sub
     Call ControlaFuncionamientoTriggerSaldosLotes(OPX% - 1)
   Else
     Call COMUNI("Proceso Seleccionado No Posee Control")
   End If
End Sub

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(EMPREAC$)
    VERANTER$ = "CONSTOCK"
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
  
    Call CARGA_ENCABEZADO(TRIGCONTROL.MSF, "Nombre/A24;Creado/D16;Modificado/D16;RutinaSP/A0", TRIGCONTROL, 0, 0)
''    Call GENERAR_SALDOS_X_CODIGO_LOTE_Y_DEPOSITO

End Sub

Private Sub GRID_Click()
    Static MODO  As Boolean
    If (GRID.MouseRow = 0) Then
          
        ' Ordena en forma ascendente
        If MODO Then
            GRID.COL = GRID.MouseCol
            GRID.Sort = 2
            MODO = False
        ' Ordena en forma descendente
        Else
            GRID.COL = GRID.MouseCol
            GRID.Sort = 1
            MODO = True
        End If
    End If
End Sub

Private Sub GRID_DblClick()
   REG& = GRID.MouseRow
   If REG& < 1 Then Exit Sub
   If GRID.Rows < 2 Then Exit Sub
   CODIX$ = GRID.TextMatrix(REG&, 1)
   CI1% = CODINT(CODIX$)
   If CI1% > 0 Then
     FIKARDEX09.Text1 = CODIX$
     FIKARDEX09.Show 1
   End If
End Sub


Private Sub mnuControlSalotedep1_Click()
       Screen.MousePointer = 11
       Call ABRESTOCKS
       Call CARGA_ENCABEZADO(TRIGCONTROL.GRID, "Código/A18;Descripción/a24;Dep./F4;Lote/A12;SaldMov/F10.1;Saldo N.T./F10.1;Difer./F10.1;", TRIGCONTROL, 1)
       Call CARGA_ENCABEZADO(TRIGCONTROL.MSF_2, "Código/A18;Descripción/a24;Dep./F4;Lote/A12;SaldMov/F10.1;Saldo N.T./F10.1;Difer./F10.1;", TRIGCONTROL, 1)
       '
       Call PoneEnCeroNull("MOVISTO", "CANTINGRE", REGAF&, "NOMESS")
       Call PoneEnCeroNull("MOVISTO", "CANTSALID", REGAF&, "NOMESS")
       '
       J& = 0: VUELTA% = 0: PASO% = 0
       SQLX$ = "SELECT SUM(CANTINGRE) AS INGRE,SUM(CANTSALID)AS SALI,COD_INTE, COUNT(*) AS CANTREG"
       If MODO% > 0 Then SQLX$ = SQLX$ + " ,DEPOMOVI "
       If MODO% > 1 Then SQLX$ = SQLX$ + " ,IDENLOTE"
       SQLX$ = SQLX$ + " FROM MOVISTO WITH(NOLOCK)"
       SQLX$ = SQLX$ + " WHERE COD_INTE > 0  "
       SQLX$ = SQLX$ + " AND CODIEMPR =  " & CodiEmp%
       SQLX$ = SQLX$ + " GROUP BY COD_INTE"
       If MODO% > 0 Then SQLX$ = SQLX$ + " ,DEPOMOVI"
       If MODO% > 1 Then SQLX$ = SQLX$ + " ,IDENLOTE"
       SQLX$ = SQLX$ + " ,CODIEMPR"
       SQLX$ = SQLX$ + " ORDER BY COD_INTE"
       Debug.Print SQLX$
       Set rs = READSET(SQLX$)
       With rs
        Do While Not .EOF
            If PASO% = 0 Then
               FIN& = XVALO(!CANTREG)
               PASO% = 1
            End If
            IJ& = IJ& + 1
            Call TRACE(20, IJ&, FIN&)
            CI2& = XVALO(!cod_inte): CI1% = CI2&
            SALDO = XVALO(!INGRE) - XVALO(!sali)
            If MODO% > 0 Then DEPO% = XVALO(!depomovi)
            If MODO% > 1 Then LOTE$ = SAFETEXT$(!idenlote)
            '
            'LEO EL SALDO DE LA TABLA SALOTEDEP1
            CONDI$ = "COD_INTE = " & CI2&
            If MODO% > 0 Then CONDI$ = CONDI$ + " AND DEPOMOVI = " & DEPO%
            If MODO% > 1 Then CONDI$ = CONDI$ + " AND IDENLOTE = '" & LOTE$ & "'"
            CONDI$ = CONDI$ + " AND CODIEMPR =  " & CodiEmp%
            CONDI$ = CONDI$ + " GROUP BY COD_INTE"
            If MODO% > 0 Then CONDI$ = CONDI$ + " ,DEPOMOVI"
            If MODO% > 1 Then CONDI$ = CONDI$ + " ,IDENLOTE"
            CONDI$ = CONDI$ + " ,CODIEMPR"
            SALDO_TN = XVALO(VALOBADA("SALOTEDEP1", "SUM(SALDO)", CONDI$, "NOMESS"))
            '
            DIFERENCIAX = SALDO - SALDO_TN
            If Abs(DIFERENCIAX) > 0 Then
               J& = J& + 1
               GRID.Rows = J& + 1
               GRID.TextMatrix(J&, 1) = CODEXT$(CI1%)
               GRID.TextMatrix(J&, 2) = DESCRIT0$(CI1%)
               If MODO% > 0 Then GRID.TextMatrix(J&, 4) = LOTE$
               If MODO% > 1 Then GRID.TextMatrix(J&, 3) = DEPO%
               GRID.TextMatrix(J&, 5) = FORMATNUM$(SALDO, "F10.1")
               GRID.TextMatrix(J&, 6) = FORMATNUM$(SALDO_TN, "F10.1")
               GRID.TextMatrix(J&, 7) = FORMATNUM$(DIFERENCIAX, "F10.1")
            End If
           .MoveNext
        Loop
       End With
      rs.Close
      Set rs = Nothing
      Screen.MousePointer = 1
      If J& < 1 Then
        Call COMUNI("No se Detectaron Diferencias de Saldo Entre el Movisto y La Nueva Tabla SALOTEDEP1")
      Else
        Screen.MousePointer = 1
        Call COMUNI("Proceso Terminado")
      End If
      
End Sub


Private Sub mnuCopiaSalotedepASalotedep1_Click()
   'ESTO INICIA UN PROCESO PARA COPIAR LA TABLA SALOTEDEP A SALOTEDEP1
   'ESTO ES EL PASO PREVIO A REALIZAR LA GENERACION O ACTUALIZACION DEL MOVISTO SOBRE LA TABLA SALOTEDEP
   'LA IDEA ES
   '1.-MIGRAR LA INFORMACION DE MOVISTO A SALOTEDEP
   '2.-VERIFICAR QUE NO HAYA DIFERENCIAS ENTTRE MOVISTO Y SALOTEDEP
   '3.-CREAR LA TABLA SALOTEDEP1
   '4.-CORRER PROCESO DE ASIGNAR LOTES A MOVIMIENTOS
    'a) RecuperarLote (ASIGNAR LOTES A MOVIMIENTOS RT SIN LOTE)
    'b) RecuperarLoteNotaDeCreditoParciales (ASIGNAR LOTES TOTALES O PARCIALES A NOTAS DE CREDITO )
    'c) RecuperarLote (ASIGNAR LOTES A MOVIMIENTOS RT SIN LOTE)
   '5.-Correr proceso para comparar los saldos de stocks actuales del movisto contra la tabla salotedep1
   '6.-SI HAY DIFERENCIAS SE DEBERA VOLVER ATRAS EL BACKUP
   
    'CREO LA TABLA SALOTEDEP1
    Call CREACAMPO("STOCKS", "SALOTEDEP1", "COD_INTE", 4, 0, 1)
    Call CREACAMPO("STOCKS", "SALOTEDEP1", "CODIEMPR", 3, 0)
    Call CREACAMPO("STOCKS", "SALOTEDEP1", "IdenLote", 10, 16)
    Call CREACAMPO("STOCKS", "SALOTEDEP1", "SALDO", 7, 0)
    Call CREACAMPO("STOCKS", "SALOTEDEP1", "DepoMovi", 3, 0, 1)
    Call CREACAMPO("STOCKS", "SALOTEDEP1", "FechMov", 8, 0, 1)
    Call CREACAMPO("STOCKS", "SALOTEDEP1", "Nume_Cli", 4, 0, 1)
    '
    Call ABRESTOCKS
    
    If CLACONTRA% > 0 Then
       Screen.MousePointer = 11
       Call PoneEnCeroNull("MOVISTO", "CANTINGRE", REGAF&, "NOMESS")
       Call PoneEnCeroNull("MOVISTO", "CANTSALID", REGAF&, "NOMESS")
       
       SQLX$ = "DELETE FROM SALOTEDEP1"
       FLEXCONN.Execute SQLX$
       
       SQLX$ = "SELECT SUM(CANTINGRE) AS INGRE,SUM(CANTSALID)AS SALI,COD_INTE,IDENLOTE,DEPOMOVI,CODIEMPR FROM MOVISTO WITH(NOLOCK)"
       SQLX$ = SQLX$ + " WHERE COD_INTE > 0  and CODIEMPR =  " & CodiEmp%
       If CI1% > 0 Then SQLX$ = SQLX$ + " AND COD_INTE = " & CI1%
       SQLX$ = SQLX$ + " GROUP BY COD_INTE,IDENLOTE,DEPOMOVI,CODIEMPR"
       SQLX$ = SQLX$ + " ORDER BY COD_INTE,IDENLOTE"
       Set rs = READSET(SQLX$)
       With rs
        Do While Not .EOF
            CI2& = XVALO(!cod_inte)
            SALDO = XVALO(!INGRE) - XVALO(!sali)
            LOTE$ = SAFETEXT$(!idenlote)
            DEPO% = XVALO(!depomovi)
            FECHALTAX = FETEXCOR1$(FECREALOT%(LOTE$))
            
            NPROV = PROVELOTE%(LOTE$)
            'AGREGA EL REGISTRO
            SQLX2$ = "INSERT INTO SALOTEDEP1"
            SQLX2$ = SQLX2$ + " (COD_INTE,IdenLote,DepoMovi,CODIEMPR,SALDO,FechMov,Nume_Cli)"
            SQLX2$ = SQLX2$ + " VALUES(" & CI2&
            SQLX2$ = SQLX2$ + " ,'" & LOTE$ & "'"
            SQLX2$ = SQLX2$ + " ," & DEPO%
            SQLX2$ = SQLX2$ + " ," & CodiEmp%
            SQLX2$ = SQLX2$ + " ," & SALDO
            SQLX2$ = SQLX2$ + " ,'" & FECHALTAX & "'"
            SQLX2$ = SQLX2$ + " ," & NPROV
            SQLX2$ = SQLX2$ + ")"
            FLEXCONN.Execute (SQLX2$)
          .MoveNext
        Loop
       End With
    
    
      Call COMUNI("Proceso Realizado Exitosamente")
    End If
 
   
   

End Sub

Private Sub mnumovistosalotedep_Click()
        Call COMUNI("Atención !!! Si Acepta Tendrá La Ventana de Selección de Código\Si Selecciona Código, Se Visualizará Solo Los Saldos de Dicho Código\Para Todos Los Códigos\No Seleccione y Presione Aceptar")
10      OBJ3$ = OBJPLA$("PRESCODIGO", VDEF$)
        
        If OBJ3$ <> "" Then
            CI1% = CVI(Mid$(OBJ3$, 1, 2))
            'VALIDACIONES
            If CI1% <= 0 Then
              OPX% = COMALTER%("No se Ha ingresado Código o El Mismo es Inexistente\ Si Selecciona Ver Informe \Se cargará la Información de Todos Los Códigos\\Cancelar\Ver Todos")
              If OPX% < 2 Then Exit Sub
            End If
       Else
            Exit Sub
       End If
       OPX% = COMALTER%("Seleccionar Controles\\Control de Saldos X Items\Control de Saldos X Items-Depósitos\Control de Saldos X Items-Lote-Deposito")
       If OPX% < 1 Then Exit Sub
       MODO% = OPX% - 1

       Screen.MousePointer = 11
       Call PoneEnCeroNull("MOVISTO", "CANTINGRE", REGAF&, "NOMESS")
       Call PoneEnCeroNull("MOVISTO", "CANTSALID", REGAF&, "NOMESS")

       Call CARGA_ENCABEZADO(TRIGCONTROL.GRID, "Código/A18;Descripción/a24;Dep./F4;Lote/A12;SaldMov/F10.1;Saldo N.T./F10.1;Difer./F10.1;", TRIGCONTROL, 1)
       Call CARGA_ENCABEZADO(TRIGCONTROL.MSF_2, "Código/A18;Descripción/a24;Dep./F4;Lote/A12;SaldMov/F10.1;Saldo N.T./F10.1;Difer./F10.1;", TRIGCONTROL, 1)
       '
       J& = 0: VUELTA% = 0: PASO% = 0
       SQLX$ = "SELECT SUM(CANTINGRE) AS INGRE,SUM(CANTSALID)AS SALI,COD_INTE, COUNT(*) AS CANTREG"
       If MODO% > 0 Then SQLX$ = SQLX$ + " ,DEPOMOVI "
       If MODO% > 1 Then SQLX$ = SQLX$ + " ,IDENLOTE"
       SQLX$ = SQLX$ + " FROM MOVISTO WITH(NOLOCK)"
       SQLX$ = SQLX$ + " WHERE COD_INTE > 0  "
       If CI1% > 0 Then SQLX$ = SQLX$ + " AND COD_INTE = " & CI1%
       SQLX$ = SQLX$ + " AND CODIEMPR =  " & CodiEmp%
       SQLX$ = SQLX$ + " GROUP BY COD_INTE"
       If MODO% > 0 Then SQLX$ = SQLX$ + " ,DEPOMOVI"
       If MODO% > 1 Then SQLX$ = SQLX$ + " ,IDENLOTE"
       SQLX$ = SQLX$ + " ,CODIEMPR"
       SQLX$ = SQLX$ + " ORDER BY COD_INTE"
       Debug.Print SQLX$
       Set rs = READSET(SQLX$)
       With rs
        Do While Not .EOF
            If PASO% = 0 Then
               FIN& = XVALO(!CANTREG)
               PASO% = 1
            End If
            IJ& = IJ& + 1
            Call TRACE(20, IJ&, FIN&)
            CI2& = XVALO(!cod_inte): CI1% = CI2&
            SALDO = XVALO(!INGRE) - XVALO(!sali)
            If MODO% > 0 Then DEPO% = XVALO(!depomovi)
            If MODO% > 1 Then LOTE$ = SAFETEXT$(!idenlote)
            '
            'LEO EL SALDO DE LA TABLA SALOTEDEP
            CONDI$ = "COD_INTE = " & CI2&
            If MODO% > 0 Then CONDI$ = CONDI$ + " AND DEPOMOVI = " & DEPO%
            If MODO% > 1 Then CONDI$ = CONDI$ + " AND IDENLOTE = '" & LOTE$ & "'"
            CONDI$ = CONDI$ + " AND CODIEMPR =  " & CodiEmp%
            CONDI$ = CONDI$ + " GROUP BY COD_INTE"
            If MODO% > 0 Then CONDI$ = CONDI$ + " ,DEPOMOVI"
            If MODO% > 1 Then CONDI$ = CONDI$ + " ,IDENLOTE"
            CONDI$ = CONDI$ + " ,CODIEMPR"
            SALDO_TN = XVALO(VALOBADA("SALOTEDEP", "SUM(SALDO)", CONDI$, "NOMESS"))
            '
            J& = J& + 1
            GRID.Rows = J& + 1
            GRID.TextMatrix(J&, 1) = CODEXT$(CI1%)
            GRID.TextMatrix(J&, 2) = DESCRIT0$(CI1%)
            If MODO% > 0 Then GRID.TextMatrix(J&, 4) = LOTE$
            If MODO% > 1 Then GRID.TextMatrix(J&, 3) = DEPO%
            GRID.TextMatrix(J&, 5) = FORMATNUM$(SALDO, "F10.1")
            GRID.TextMatrix(J&, 6) = FORMATNUM$(SALDO_TN, "F10.1")
            GRID.TextMatrix(J&, 7) = FORMATNUM$(DIFERENCIAX, "F10.1")
            '
           .MoveNext
        Loop
       End With
      rs.Close
      Set rs = Nothing
      Screen.MousePointer = 1
      If J& < 1 Then
        Call COMUNI("No se Detectaron Diferencias de Saldo Entre el Movisto y La Nueva Tabla Salotedep")
      Else
        Screen.MousePointer = 1
        Call COMUNI("Proceso Terminado")
      End If
End Sub

Private Sub mnuvertrigger_Click()
MSF.Rows = 1
Me.TXTSTORED = ""
SQLX$ = "SELECT * FROM SYS.OBJECTS WITH(NOLOCK) WHERE TYPE = 'TR'"
Set rs = READSET(SQLX$)
With rs
  Do While Not .EOF
    J& = J& + 1
    MSF.Rows = J& + 1
    MSF.TextMatrix(J&, 1) = SAFETEXT$(!Name)
    MSF.TextMatrix(J&, 2) = Format(!CREATE_DATE, "dd/MM/yyyy H:mm")
    MSF.TextMatrix(J&, 3) = Format(!MODIFY_DATE, "dd/MM/yyyy H:mm")
    MSF.TextMatrix(J&, 4) = ""
    .MoveNext
  Loop
End With
End Sub

Private Sub mnuVisualizarSP_Click()
MSF.Rows = 1
SQLX$ = "select SPECIFIC_NAME,CREATED,LAST_ALTERED,ROUTINE_DEFINITION "
SQLX$ = SQLX$ + "  From " & BASEACTIVA$ & ".information_schema.routines WITH(NOLOCK)"
SQLX$ = SQLX$ + "  where routine_type = 'PROCEDURE'"
SQLX$ = SQLX$ + "  and  Left ( Routine_Name ,  3 )  NOT  IN  ( 'sp_' ,  'xp_' ,  'ms_' )"

Set rs = READSET(SQLX$)
With rs
  Do While Not .EOF
    J& = J& + 1
    MSF.Rows = J& + 1
    MSF.TextMatrix(J&, 1) = SAFETEXT$(!SPECIFIC_NAME)
    MSF.TextMatrix(J&, 2) = Format(!CREATED, "dd/MM/yyyy H:mm")
    MSF.TextMatrix(J&, 3) = Format(!LAST_ALTERED, "dd/MM/yyyy H:mm")
    MSF.TextMatrix(J&, 4) = SAFETEXT$(!ROUTINE_DEFINITION)
    .MoveNext
  Loop
End With
End Sub


Private Sub MSF_2_Click()
    Static MODO  As Boolean
    If (MSF_2.MouseRow = 0) Then
          
        ' Ordena en forma ascendente
        If MODO Then
            MSF_2.COL = MSF_2.MouseCol
            MSF_2.Sort = 2
            MODO = False
        ' Ordena en forma descendente
        Else
            MSF_2.COL = MSF_2.MouseCol
            MSF_2.Sort = 1
            MODO = True
        End If
    End If
End Sub



Private Sub MSF2_Click()
    Static MODO  As Boolean
    If (MSF2.MouseRow = 0) Then
          
        ' Ordena en forma ascendente
        If MODO Then
            MSF2.COL = MSF2.MouseCol
            MSF2.Sort = 2
            MODO = False
        ' Ordena en forma descendente
        Else
            MSF2.COL = MSF2.MouseCol
            MSF2.Sort = 1
            MODO = True
        End If
    End If

End Sub




Private Sub MSF_Click()
    TXTSTORED = ""
    REG& = XVALO(MSF.MouseRow)
    MSF.TextMatrix(REG&, 0) = "*" 'MARCA DE ASTERIZCO AL SELECCIONADO
    NAMETRIGGER$ = TRIM$(MSF.TextMatrix(REG&, 1))
    SQLX$ = "SELECT  m.definition AS TEXTO"
    '+ char(13) +'Go'"
    '
    For i& = 1 To MSF.Rows - 1
      If i& <> REG& Then MSF.TextMatrix(i&, 0) = "" 'LIMPIO TODOS MENOS EL MARCADO ULTIMO
    Next i&
    '
    'STORED PROCEDURE
    '****************
    'SI ES STORED PROCEDURE GUARDA EL CODIGO DEL MISMO EN LA COLUMNA 4 (NO VISIBLE)
    'MUESTRA Y NO SIGUE
    If TRIM$(MSF.TextMatrix(REG&, 4)) <> "" Then
        TXTSTORED = MSF.TextMatrix(REG&, 4)
        Exit Sub
    End If
    '
    'TRIGGER
    '*******
    SQLX$ = SQLX$ + " FROM sys.sql_modules as M WITH(NOLOCK) inner join sys.objects  as O"
    SQLX$ = SQLX$ + " on"
    SQLX$ = SQLX$ + " M.object_id = O.object_id"
    SQLX$ = SQLX$ + "  Where (O.Type = 'TR'  OR O.Type = 'SP' ) AND M.DEFINITION LIKE '%" & NAMETRIGGER$ & "%'"
    Set rs = READSET(SQLX$)
    With rs
     If Not .EOF Then TXTSTORED = SAFETEXT$(!Texto)
    End With
End Sub

Private Sub TXTBUSCAR_Change()
    '
    If TRIM$(TXTBUSCAR) = "" Then
      MSF_2.Visible = False
      GRID.Visible = True
      Exit Sub
    End If
    '
    MSF_2.Rows = 2
    For i& = 1 To MSF_2.Cols - 1
       MSF_2.TextMatrix(1, i&) = ""
    Next i&
    REG& = 0
    MSF_2.Visible = True
    MSF_2.ZOrder 0
    MSF_2.Cols = GRID.Cols
    MSF_2.Redraw = False
    GRID.Visible = False
    '
    'RECORRO EL MSF QUE TIENE TODO E IMPRIMO EN msf  LOS QUE MACHEAN
    For i& = 1 To GRID.Rows - 1
      For J& = 1 To GRID.Cols - 1
        If (InStr(UCase$(GRID.TextMatrix(i&, J&)), UCase$(TXTBUSCAR))) Then
          If (TRIM$(CMBFILTRO.TEXT)) = "" Or TRIM$(UCase$(CMBFILTRO)) = TRIM$(UCase$(GRID.TextMatrix(0, J&))) Then
            REG& = REG& + 1
            MSF_2.Rows = REG& + 1
            For H& = 1 To MSF_2.Cols - 1
              MSF_2.TextMatrix(REG&, H&) = GRID.TextMatrix(i&, H&)
            Next H&
          End If
          If (TRIM$(CMBFILTRO.TEXT)) = "" Then Exit For
        End If
      Next J&
    Next i&
    '
    MSF_2.Redraw = True
    
    '
End Sub

Private Sub TXTBUSCAR_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      TXTBUSCAR.TEXT = ""
   End If
End Sub
Sub ControlaFuncionamientoTriggerSaldosLotesAnt(MODO%)
       'MODO% = 0 AGRUPA POR CODIGO
       'MODO% = 1 AGRUPA POR CODIGO,LOTE
       'MODO% = 2 AGRUPA POR CODIGO,LOTE, DEPOSITO
       Call ABRESTOCKS
       Call CARGA_ENCABEZADO(TRIGCONTROL.GRID, "Código/A18;Descripción/a24;Dep./F4;Lote/A12;SaldMov/F10.1;Saldo N.T./F10.1;Difer./F10.1;", TRIGCONTROL, 1)
       Call CARGA_ENCABEZADO(TRIGCONTROL.MSF_2, "Código/A18;Descripción/a24;Dep./F4;Lote/A12;SaldMov/F10.1;Saldo N.T./F10.1;Difer./F10.1;", TRIGCONTROL, 1)
       '
       Call PoneEnCeroNull("MOVISTO", "CANTINGRE", REGAF&, "NOMESS")
       Call PoneEnCeroNull("MOVISTO", "CANTSALID", REGAF&, "NOMESS")
       '
       J& = 0: VUELTA% = 0
       SQLX$ = "SELECT SUM(CANTINGRE) AS INGRE,SUM(CANTSALID)AS SALI,COD_INTE,DEPOMOVI,IDENLOTE,CODIEMPR"
'       If MODO% > 0 Then SQLX$ = SQLX$ + " ,DEPOMOVI "
'       If MODO% > 1 Then SQLX$ = SQLX$ + " ,IDENLOTE"
       SQLX$ = SQLX$ + " FROM MOVISTO WITH(NOLOCK)"
       SQLX$ = SQLX$ + " WHERE COD_INTE > 0  "
       SQLX$ = SQLX$ + " GROUP BY COD_INTE,DEPOMOVI,IDENLOTE,CODIEMPR"
'       If MODO% > 0 Then SQLX$ = SQLX$ + " ,DEPOMOVI"
'       If MODO% > 1 Then SQLX$ = SQLX$ + " ,IDENLOTE"
       SQLX$ = SQLX$ + " ORDER BY COD_INTE"
       Set rs = READSET(SQLX$)
       With rs
        Do While Not .EOF
            CI2& = XVALO(!cod_inte): CI1% = CI2&
            SALDO = XVALO(!INGRE) - XVALO(!sali)
            DEPO% = XVALO(!depomovi)
            LOTE$ = SAFETEXT$(!idenlote)
'            If MODO% > 0 Then DEPO% = XVALO(!DEPOMOVI)
'            If MODO% > 1 Then LOTE$ = SAFETEXT$(!IDENLOTE)
            '
            'LEO EL SALDO DE LA TABLA SALOTEDEP
            CONDI$ = "COD_INTE = " & CI2&
'            CONDI$ = CONDI$ + " AND DEPOMOVI = " & DEPO%
'            CONDI$ = CONDI$ + " AND IDENLOTE = '" & LOTE$ & "'"
            If MODO% > 0 Then CONDI$ = CONDI$ + " AND DEPOMOVI = " & DEPO%
            If MODO% > 1 Then CONDI$ = CONDI$ + " AND IDENLOTE = '" & LOTE$ & "'"
            CONDI$ = CONDI$ + " AND CODIEMPR =  " & CodiEmp%
            CONDI$ = CONDI$ + " GROUP BY COD_INTE"
'            CONDI$ = CONDI$ + " ,DEPOMOVI"
'            CONDI$ = CONDI$ + " ,IDENLOTE"
            CONDI$ = CONDI$ + " ,CODIEMPR"
            If MODO% > 0 Then CONDI$ = CONDI$ + " ,DEPOMOVI"
            If MODO% > 1 Then CONDI$ = CONDI$ + " ,IDENLOTE"
            
            SALDO_TN = XVALO(VALOBADA("SALOTEDEP", "SUM(SALDO)", CONDI$, "NOMESS"))
            '
            DIFERENCIAX = SALDO - SALDO_TN
            If Abs(DIFERENCIAX) > 0 Then
               If VUELTA% < 1 Then ' PARA QUE MUESTRE LA CONSULTA EN LA CAJA DE TEXTO
                 VUELTA% = 1
                 Me.TXTSTORED = ""
                 Me.TXTSTORED = "SELECT   SUM(SALDO) FROM SALOTEDEP " & vbCrLf
                 Me.TXTSTORED = Me.TXTSTORED & "WHERE " & CONDI$ & vbCrLf
                 Me.TXTSTORED = Me.TXTSTORED & "___________________________" & vbCrLf
                 Me.TXTSTORED = Me.TXTSTORED & SQLX$
               End If
               J& = J& + 1
               GRID.Rows = J& + 1
               GRID.TextMatrix(J&, 1) = CODEXT$(CI1%)
               GRID.TextMatrix(J&, 2) = DESCRIT0$(CI1%)
               GRID.TextMatrix(J&, 3) = DEPO%
               GRID.TextMatrix(J&, 4) = LOTE$
'               If MODO% > 0 Then GRID.TextMatrix(j&, 4) = LOTE$
'               If MODO% > 1 Then GRID.TextMatrix(j&, 3) = DEPO%
               GRID.TextMatrix(J&, 5) = FORMATNUM$(SALDO, "F10.1")
               GRID.TextMatrix(J&, 6) = FORMATNUM$(SALDO_TN, "F10.1")
               GRID.TextMatrix(J&, 7) = FORMATNUM$(DIFERENCIAX, "F10.1")
            End If
           .MoveNext
        Loop
       End With
      rs.Close
      Set rs = Nothing
      If J& < 1 Then
        Call COMUNI("No se Detectaron Diferencias de Saldo Entre el Movisto y La Nueva Tabla Salotedep")
      Else
        Call COMUNI("Proceso Terminado")
      End If
      
End Sub

Sub ControlaFuncionamientoTriggerSaldosLotes(MODO%)
       'MODO% = 0 AGRUPA POR CODIGO
       'MODO% = 1 AGRUPA POR CODIGO,LOTE
       'MODO% = 2 AGRUPA POR CODIGO,LOTE, DEPOSITO
       Screen.MousePointer = 11
       Call ABRESTOCKS
       Call CARGA_ENCABEZADO(TRIGCONTROL.GRID, "Código/A18;Descripción/a24;Dep./F4;Lote/A12;SaldMov/F10.1;Saldo N.T./F10.1;Difer./F10.1;", TRIGCONTROL, 1)
       Call CARGA_ENCABEZADO(TRIGCONTROL.MSF_2, "Código/A18;Descripción/a24;Dep./F4;Lote/A12;SaldMov/F10.1;Saldo N.T./F10.1;Difer./F10.1;", TRIGCONTROL, 1)
       '
       Call PoneEnCeroNull("MOVISTO", "CANTINGRE", REGAF&, "NOMESS")
       Call PoneEnCeroNull("MOVISTO", "CANTSALID", REGAF&, "NOMESS")
       '
       J& = 0: VUELTA% = 0: PASO% = 0
       SQLX$ = "SELECT SUM(CANTINGRE) AS INGRE,SUM(CANTSALID)AS SALI,COD_INTE, COUNT(*) AS CANTREG"
       If MODO% > 0 Then SQLX$ = SQLX$ + " ,DEPOMOVI "
       If MODO% > 1 Then SQLX$ = SQLX$ + " ,IDENLOTE"
       SQLX$ = SQLX$ + " FROM MOVISTO WITH(NOLOCK)"
       SQLX$ = SQLX$ + " WHERE COD_INTE > 0  "
       SQLX$ = SQLX$ + " AND CODIEMPR =  " & CodiEmp%
       SQLX$ = SQLX$ + " GROUP BY COD_INTE"
       If MODO% > 0 Then SQLX$ = SQLX$ + " ,DEPOMOVI"
       If MODO% > 1 Then SQLX$ = SQLX$ + " ,IDENLOTE"
       SQLX$ = SQLX$ + " ,CODIEMPR"
       SQLX$ = SQLX$ + " ORDER BY COD_INTE"
'       Debug.Print SQLX$
       Set rs = READSET(SQLX$)
       With rs
        Do While Not .EOF
            If PASO% = 0 Then
               FIN& = XVALO(!CANTREG)
               PASO% = 1
            End If
            IJ& = IJ& + 1
            Call TRACE(20, IJ&, FIN&)
            CI2& = XVALO(!cod_inte): CI1% = CI2&
            SALDO = XVALO(!INGRE) - XVALO(!sali)
            If MODO% > 0 Then DEPO% = XVALO(!depomovi)
            If MODO% > 1 Then LOTE$ = SAFETEXT$(!idenlote)
            '
            'LEO EL SALDO DE LA TABLA SALOTEDEP
            CONDI$ = "COD_INTE = " & CI2&
            If MODO% > 0 Then CONDI$ = CONDI$ + " AND DEPOMOVI = " & DEPO%
            If MODO% > 1 Then CONDI$ = CONDI$ + " AND IDENLOTE = '" & LOTE$ & "'"
            CONDI$ = CONDI$ + " AND CODIEMPR =  " & CodiEmp%
            CONDI$ = CONDI$ + " GROUP BY COD_INTE"
            If MODO% > 0 Then CONDI$ = CONDI$ + " ,DEPOMOVI"
            If MODO% > 1 Then CONDI$ = CONDI$ + " ,IDENLOTE"
            CONDI$ = CONDI$ + " ,CODIEMPR"
            SALDO_TN = XVALO(VALOBADA("SALOTEDEP", "SUM(SALDO)", CONDI$, "NOMESS"))
            '
            Debug.Print CONDI$
            DIFERENCIAX = SALDO - SALDO_TN
            If Abs(DIFERENCIAX) > 0 Then
               If VUELTA% < 1 Then ' PARA QUE MUESTRE LA CONSULTA EN LA CAJA DE TEXTO
                 VUELTA% = 1
                 Me.TXTSTORED = ""
                 Me.TXTSTORED = "SELECT   SUM(SALDO) FROM SALOTEDEP " & vbCrLf
                 Me.TXTSTORED = Me.TXTSTORED & "WHERE " & CONDI$ & vbCrLf
                 Me.TXTSTORED = Me.TXTSTORED & "___________________________" & vbCrLf
                 Me.TXTSTORED = Me.TXTSTORED & SQLX$
               End If
               J& = J& + 1
               GRID.Rows = J& + 1
               GRID.TextMatrix(J&, 1) = CODEXT$(CI1%)
               GRID.TextMatrix(J&, 2) = DESCRIT0$(CI1%)
               If MODO% > 0 Then GRID.TextMatrix(J&, 4) = LOTE$
               If MODO% > 1 Then GRID.TextMatrix(J&, 3) = DEPO%
               GRID.TextMatrix(J&, 5) = FORMATNUM$(SALDO, "F10.1")
               GRID.TextMatrix(J&, 6) = FORMATNUM$(SALDO_TN, "F10.1")
               GRID.TextMatrix(J&, 7) = FORMATNUM$(DIFERENCIAX, "F10.1")
            End If
           .MoveNext
        Loop
       End With
      rs.Close
      Set rs = Nothing
      Screen.MousePointer = 1
      If J& < 1 Then
        Call COMUNI("No se Detectaron Diferencias de Saldo Entre el Movisto y La Nueva Tabla Salotedep")
      Else
        Screen.MousePointer = 1
        Call COMUNI("Proceso Terminado")
      End If
      
End Sub


