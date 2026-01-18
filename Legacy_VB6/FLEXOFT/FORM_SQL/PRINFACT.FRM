VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form PRINFACT 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   12735
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   12735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
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
      Left            =   11040
      TabIndex        =   7
      Top             =   7320
      Visible         =   0   'False
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
      OleObjectBlob   =   "PRINFACT.frx":0000
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   120
      OleObjectBlob   =   "PRINFACT.frx":0234
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1485
   End
   Begin MSFlexGridLib.MSFlexGrid GRID 
      Height          =   7125
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Un Click Sobre Encabezado Para Ordenar"
      Top             =   120
      Width           =   12555
      _ExtentX        =   22146
      _ExtentY        =   12568
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
      Height          =   7125
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Un Click Sobre Encabezado Para Ordenar"
      Top             =   120
      Visible         =   0   'False
      Width           =   12555
      _ExtentX        =   22146
      _ExtentY        =   12568
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
      Begin VB.Menu MnuRevertirSeleccion 
         Caption         =   "Revertir Selección"
      End
      Begin VB.Menu mnuimprimir 
         Caption         =   "Imprimir Documentos Seleccionados"
      End
      Begin VB.Menu mnucorrelatividad 
         Caption         =   "Detectar Falta de Correlatividad"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "PRINFACT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub BOTEXIT_Click()
  Unload Me
End Sub

Private Sub CMBFILTRO_Change()
   TXTBUSCAR_Change
End Sub

Private Sub Command1_Click()
   Command1.Enabled = False
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

99 Command1.Enabled = True

End Sub
Sub GENEXCEL(NOMBREDEARCHIVO$, GRID As MSFlexGrid)
   Dim XLAPP As Excel.Application
   Dim XLBOOK As Excel.Workbook
   Dim XLSHEET As Excel.Worksheet
   '
   Dim FECHA As Date
   Dim valor As Double
   Dim TEXTO As String
   '
   On Error GoTo 20
   Set XLAPP = New Excel.Application
   Set XLBOOK = XLAPP.Workbooks.Add
   Set XLSHEET = XLBOOK.Worksheets.Add
   On Error GoTo 0
   '
   TIPODATO% = 0: COLUM& = 0
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
             TIPODATO% = 1
          End If
          ' ESTO ES PARA QUE EXPORTE COMO CERO Y NO ESPACIOS
          FFCC$ = UCase$(Left$(FORCAMPO$(COLU1&), 1))
          If FFCC$ = "I" Or FFCC$ = "F" Or FFCC$ = "G" Or FFCC$ = "$" Then
             If TRIM$(CStr(VALORX)) = "" Then
                VALORX = 0
             End If
             TIPODATO% = 2
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
          If TIPODATO% = 1 Then XLSHEET.Cells(RENG1&, COLUM&).NumberFormat = "dd/mm/yyyy"
          'If TIPODATO% = 2 Then XLSHEET.Cells(RENG1&, COLU1&).NumberFormat = "#,##0.00"
          If TIPODATO% = 0 Then XLSHEET.Cells(RENG1&, COLUM&).NumberFormat = "@"
          TIPODATO% = 0
          VALORX_ANT = VALORX
10     Next COLU1&
   Next i&
   '
   Close #n1%
   On Error GoTo 30
   XLSHEET.SaveAs NOMBREDEARCHIVO$
   On Error GoTo 0
   '
   XLAPP.Quit      ' cierra EXCEL que se esta ejecutando minimizado
   Set XLAPP = Nothing     ') libera memoria
   Set XLBOOK = Nothing    ') utilizada por
   Set XLSHEET = Nothing   ') las variables objeto
   '
   Screen.MousePointer = 1
   Call COMUNI("Planilla Excel Creada Exitosamente:\ \" + REPLACAR$(NOMBREDEARCHIVO$, "\", "/"))
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

Private Sub Label13_Change()
End Sub

Private Sub Label9_Change()
    
End Sub



Private Sub Command2_Click()

End Sub

Private Sub GRID_Click()
    Static MODO  As Boolean
    If (GRID.MouseRow = 0) Then
        PASOFECHA% = 0
        ' Ordena en forma ascendente
        GRID.COL = GRID.MouseCol
        If InStr(UCase(GRID.TextMatrix(0, GRID.COL)), "FECH") > 0 Then
           Call ORDENARCOLUFECHA(GRID, GRID.COL)
           PASOFECHA% = 1
        End If

        If MODO Then
            GRID.Sort = 2
            MODO = False
        ' Ordena en forma descendente
        Else
            GRID.COL = GRID.MouseCol
            GRID.Sort = 1
            MODO = True
        End If
        If PASOFECHA% > 0 Then
           Call PASARTEXTOAFECHA(GRID, GRID.COL)
        End If
    End If
End Sub

Private Sub GRID_DblClick()
   REG& = GRID.MouseRow
   Call COLOREAR_GRILLA_SOLAFILA(GRID, &HFFFF80, REG&, 0, 0, 0)
End Sub

Private Sub mnucorrelatividad_Click()
        Dim MSFACTIVO As MSFlexGrid
        If GRID.Visible = True Then
          
        Else
          Call COMUNI("Solo sin Filtro")
          Exit Sub
        End If
        '
        'ACA SELECCIONAR LOS TIPOS DE FACTURAS
        'PRESENTAR LAS FACTURAS SELECCIONADAS
        'ORDENAR Y FILTRAR.
        
        GRID.COL = 8
        GRID.Sort = 4

        For i& = 1 To GRID.Rows - 1
           NUMEX = XVALO(GRID.TextMatrix(i&, 7))
           If NUMEX < 0 Then GoTo 50
           TIPODO$ = GRID.TextMatrix(i&, 5) 'FC-0004A
'           TIPODO$ = TIPODO$ & Right$(GRID.TextMatrix(I&, 5), 1) ' LA LETRA FINAL
           If TIPODO$ <> TIPODO_ANT$ Then
             NUMEXSIG = NUMEX
           End If
           If NUMEXSIG <> NUMEX Then
                REG& = i&
                Call COLOREAR_GRILLA_SOLAFILA(GRID, &HFFFF80, REG&, 0, 0, 0)
20              NUMEXSIG = NUMEXSIG + 1
                If NUMEXSIG <> NUMEX Then GoTo 20
           End If
           NUMEXSIG = NUMEX + 1
           TIPODO_ANT$ = TIPODO$
50      Next i&
        
        
End Sub

Private Sub mnuimprimir_Click()
        Dim MSFACTIVO As MSFlexGrid
        If GRID.Visible = True Then
          Set MSFACTIVO = GRID
        Else
          Set MSFACTIVO = MSF_2
        End If
        '
        entro% = 0
        Screen.MousePointer = 11
        For i& = 1 To MSFACTIVO.Rows - 1
          If MSFACTIVO.TextMatrix(i&, 0) = "*" Then
           If entro% = 0 Then
              entro% = 1
              CACOPI% = XVALO(InputBox$("Cantidad de Copias", "Ingrese Cantidad de Copias", Mid$(CACOPX$, 2)))
               If CACOPI% < 1 Then Exit Sub
           End If
           '
           Call SELECCIONARFILA(MSFACTIVO, i&)

           NCLI = XVALO(MSFACTIVO.TextMatrix(i&, 2))
           NUFAC$ = TRIM$(MSFACTIVO.TextMatrix(i&, 5))
           Call IMPRIMEDOCU(NUFAC$, NCLI, CACOPI%, "PRINT")
          End If
        Next i&
        Screen.MousePointer = 1
        Call COMUNI("Se Han Imprimido los Documentos Seleccionados")

End Sub
Sub IMPRIMEDOCU(NUFACTU$, NC2, CANTICOP%, MODO$)
    'Call STDFORM(ARCSQR$, FILTX$ & ";" & ORCOP$(TOTACOPIAS% - (CANTCOPIAS%)), "PRINT", , ENVIAMAILFACTURA%)
    TIPOCOM$ = Mid$(NUFACTU$, 1, 2)
    
    Condi$ = "DOSECLAR='" & NUFACTU$ & "' AND NUME_CLIE= " & NC2 & " AND COD_INTE > 0"
    CANTREG& = XVALO(CantRegistros("MOVISTO", Condi$, "NOMESS"))
    '
    TABLA1$ = "MOVISTO"
    CANTREGREPA& = 0
    If CANTREG& < 1 Then
        'VALIDO SI ES UNA REPARACION
        CONDI1$ = "DOSECLAR='" & NUFACTU$ & "' AND NUME_CLIE= " & NC2 & ""
        CANTREGREPA& = XVALO(CantRegistros("MOVIREPA", CONDI1$, "NOMESS"))
        If CANTREGREPA& > 0 Then TABLA1$ = "MOVIREPA"
    End If
    '
    'SI EL CAMPO DE Descri_Lar ES NULL O VACIO SE LE AGREGA EL DE LA DESCRIPCION
    '****************************************************************************
    SQLX$ = "UPDATE  " & TABLA1$ & " SET Descri_Lar = DESCRIP "
    SQLX$ = SQLX$ + " WHERE (Descri_Lar IS NULL OR Descri_Lar LIKE '') "
    SQLX$ = SQLX$ + " AND DOSECLAR = '" & NUFACTU$ & "' AND NUME_CLIE=" & NC2
    FLEXCONN.Execute (FILTSQL$(SQLX$))
    '
    '******************************************
    LETRA$ = Mid$(NUFACTU$, 18, 1)
    FORIPRE$ = TRIM$(Control$("", "FORFAC-" + LETRA$))
    If CANTREG& = 0 And CANTREGREPA& = 0 Then TIPOCOM$ = "(T.Libre)"
    TIFADOC% = 2
    If Mid$(NUFACTU$, 1, 2) = "NC" Then TIFADOC% = 6
    If Mid$(NUFACTU$, 1, 2) = "ND" Then TIFADOC% = 3
    '
    If TIPOCOM$ = "(T.Libre)" Then
      If LETRA$ <> "E" Then   ' EXCLUYE FACTURAS DE EXPORTACION
        FORIPRE$ = TRIM$(Control$("", "FACTEXLI"))
        If UCase$(LETRA$) = "B" Then
          FORIPRE$ = TRIM$(Control$("", "FACTELIB"))
        End If
      End If
    End If
    '
    'FORMULARIOS DE REPARACION
    If CANTREGREPA& > 0 And TIFADOC% = 2 Then
       FORIPRE$ = "FCR" + Mid$(FORIPRE$, 4)
    End If
    '
    FORIALT$ = ""
    If TIFADOC% = 3 Or TIFADOC% = 6 Then 'SI NC O ND
      If TIFADOC% = 3 Then ' SI NOTA DE DEBITO
        If UCase$(LETRA$) = "A" Then
          FORIALT$ = Control$("", "FORNDE-A") 'ND-A
          ElseIf UCase$(LETRA$) = "B" Then
          FORIALT$ = Control$("", "FORNDE-B") 'ND-B
          
          ElseIf UCase$(LETRA$) = "E" Then
          FORIALT$ = Control$("", "FORNDE-E") 'ND-E
        End If
      End If
      '
      If TIFADOC% = 6 Then ' SI NOTA DE CREDITO
        If UCase$(LETRA$) = "A" Then
          FORIALT$ = Control$("", "FORNCR-A") 'NC-A
          ElseIf UCase$(LETRA$) = "B" Then
          FORIALT$ = Control$("", "FORNCR-B") 'NC-B
          ElseIf UCase$(LETRA$) = "E" Then
          FORIALT$ = Control$("", "FORNCR-E") 'NC-E
        End If
      End If
    End If
    If TRIM$(FORIALT$) <> "" Then
      FORIPRE$ = FORIALT$
    End If
    ARCSQR$ = FORIPRE$
    If EXISTE%(LUDAT$ + ARCSQR$ + ".SQR") = 0 Then
        Call COMUNI("Falta Archivo " & ARCSQR$ + ".SQR")
    End If
    
    '
    FILTX$ = TRIM$(NUFACTU$) & ";" & TRIM$(Str$(NC2))
    Condi$ = "DOSECLAR='" & NUFACTU$ & "' AND NUME_CLIE=" & NC2
    'SE INCORPORA PARA CONTEMPLAR A LAS REPARACIONES Y OBTENER DE MOVIREPA LA MONEDA
    TABMOV$ = "MOVISTO"
    If CANTREGREPA& > 0 Then TABMOV$ = "MOVIREPA"
    MONEMI% = XVALO(VALOBADA(TABMOV$, "MONEVAL", Condi$))
    For i& = 1 To CANTICOP%
      Call STDFORM(ARCSQR$, FILTX$, "PRINT", , 0)
    Next i&
'    Call STDFORM(ARCSQR$, FILTX$, MODO$)
    '
End Sub



Private Sub MnuRevertirSeleccion_Click()
        'SETEO EL MSF QUE ESTE VISIBLE
        Dim MSFACTIVO As MSFlexGrid
        If GRID.Visible = True Then
          Set MSFACTIVO = GRID
        Else
          Set MSFACTIVO = MSF_2
        End If
        '
        Screen.MousePointer = 11
        FIN& = MSFACTIVO.Rows - 1
        For i& = 1 To MSFACTIVO.Rows - 1
           Call TRACE(20, i&, FIN&)
           Call COLOREAR_GRILLA_SOLAFILA(MSFACTIVO, &HFFFF80, i&, 0, 0, 0)
        Next i&
        Screen.MousePointer = 1
        '

End Sub

Private Sub MSF_2_Click()
    Static MODO  As Boolean
    If (MSF_2.MouseRow = 0) Then
        PASOFECHA% = 0
        ' Ordena en forma ascendente
        MSF_2.COL = MSF_2.MouseCol
        If InStr(UCase(MSF_2.TextMatrix(0, MSF_2.COL)), "FECH") > 0 Then
           Call ORDENARCOLUFECHA(MSF_2, MSF_2.COL)
           PASOFECHA% = 1
        End If
   
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
        If PASOFECHA% > 0 Then
           Call PASARTEXTOAFECHA(MSF_2, MSF_2.COL)
        End If

    End If
End Sub
Private Sub GRID_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If Button = 2 Then
    '
    i& = GRID.MouseRow
    J& = GRID.MouseCol
    If i& = 0 Or i& > GRID.Rows Then Exit Sub
    '
    On Error Resume Next
    GRID.Row = i&
    GRID.COL = J&
    On Error Resume Next

    NPX& = TRIM$(GRID.TextMatrix(i&, 1))
    If NPX& < 1 Then
        Exit Sub
    End If

    If Me.Caption = "Pedidos" Then
        '
        NUDOCU$ = TRIM$(Str$(NPX&))
        TIDOCUM$ = "Pedido de Cliente"
        While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
        NUDOCU$ = NUDOCU$ + " "
        Screen.MousePointer = 11
        '
        'BUSCA EL DOCUMENTO EN LA TABLA SQL CON LA NUEVA FUNCION
        DOCUNU& = DOCID&(TIDOCUM$, Str$(NPX&), NUDOCU$)
'        If DOCUNU& < 1 Then
'             Screen.MousePointer = 1
'             Call COMUNI("No se Encontró el Documento Impreso\Correspondiente a la " & DOCUMENTO_ELEGIDO$ & " .\  \Si se Trata de un Pedido Pendiente de Entrega\Puede Refrescarlo con la Opción\'Modificación de Pedido en Curso'.")
'             On Error Resume Next
'             OPPED07.Command46.SetFocus
'             OPPED07.Show 1
'             If Err Then
'              On Error GoTo 0
'             End If
'             Exit Sub
'        End If
        Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, Str$(NPX&), NUDOCU$)
    Else
        TIDOCUM$ = "Nota de Pedido"
        NUDOCU$ = TRIM$(Str$(NPX&))
        DOCUNU& = 0
        '
        Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, NUDOCU$)
    End If
    Screen.MousePointer = 1
  End If
End Sub

Private Sub MSF_2_DblClick()
   REG& = MSF_2.MouseRow
   Call COLOREAR_GRILLA_SOLAFILA(MSF_2, &HFFFF80, REG&, 0, 0, 0)

End Sub

Private Sub MSF_2_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If Button = 2 Then
    '
    i& = MSF_2.MouseRow
    J& = MSF_2.MouseCol
    If i& = 0 Or i& > MSF_2.Rows Then Exit Sub
    '
    On Error Resume Next
    MSF_2.Row = i&
    MSF_2.COL = J&
    On Error Resume Next

    NPX& = TRIM$(MSF_2.TextMatrix(i&, 1))
    If NPX& < 1 Then
        Exit Sub
    End If

    If Me.Caption = "Pedidos" Then
        '
        NUDOCU$ = TRIM$(Str$(NPX&))
        TIDOCUM$ = "Pedido de Cliente"
        While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
        NUDOCU$ = NUDOCU$ + " "
        Screen.MousePointer = 11
        '
        'BUSCA EL DOCUMENTO EN LA TABLA SQL CON LA NUEVA FUNCION
        DOCUNU& = DOCID&(TIDOCUM$, Str$(NPX&), NUDOCU$)
'        If DOCUNU& < 1 Then
'             Screen.MousePointer = 1
'             Call COMUNI("No se Encontró el Documento Impreso\Correspondiente a la " & DOCUMENTO_ELEGIDO$ & " .\  \Si se Trata de un Pedido Pendiente de Entrega\Puede Refrescarlo con la Opción\'Modificación de Pedido en Curso'.")
'             On Error Resume Next
'             OPPED07.Command46.SetFocus
'             OPPED07.Show 1
'             If Err Then
'              On Error GoTo 0
'             End If
'             Exit Sub
'        End If
        Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, Str$(NPX&), NUDOCU$)
    Else
        TIDOCUM$ = "Nota de Pedido"
        NUDOCU$ = TRIM$(Str$(NPX&))
        DOCUNU& = 0
        '
        Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, NUDOCU$)
    End If
    Screen.MousePointer = 1
  End If
End Sub

Private Sub msf2_Click()
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

Private Sub MSF2_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If Button = 2 Then
    '
    i& = MSF2.MouseRow
    J& = MSF2.MouseCol
    If i& = 0 Or i& > MSF2.Rows Then Exit Sub
    '
    On Error Resume Next
    MSF2.Row = i&
    MSF2.COL = J&
    On Error Resume Next

    NPX& = TRIM$(MSF2.TextMatrix(i&, 1))
    If NPX& < 1 Then
        Exit Sub
    End If

    If Me.Caption = "Pedidos" Then
        '
        NUDOCU$ = TRIM$(Str$(NPX&))
        TIDOCUM$ = "Pedido de Cliente"
        While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
        NUDOCU$ = NUDOCU$ + " "
        Screen.MousePointer = 11
        '
        'BUSCA EL DOCUMENTO EN LA TABLA SQL CON LA NUEVA FUNCION
        DOCUNU& = DOCID&(TIDOCUM$, Str$(NPX&), NUDOCU$)
'        If DOCUNU& < 1 Then
'             Screen.MousePointer = 1
'             Call COMUNI("No se Encontró el Documento Impreso\Correspondiente a la " & DOCUMENTO_ELEGIDO$ & " .\  \Si se Trata de un Pedido Pendiente de Entrega\Puede Refrescarlo con la Opción\'Modificación de Pedido en Curso'.")
'             On Error Resume Next
'             OPPED07.Command46.SetFocus
'             OPPED07.Show 1
'             If Err Then
'              On Error GoTo 0
'             End If
'             Exit Sub
'        End If
        Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, Str$(NPX&), NUDOCU$)
    Else
        TIDOCUM$ = "Nota de Pedido"
        NUDOCU$ = TRIM$(Str$(NPX&))
        DOCUNU& = 0
        '
        Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, NUDOCU$)
    End If
    Screen.MousePointer = 1
  End If
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
            For h& = 1 To MSF_2.Cols - 1
              MSF_2.TextMatrix(REG&, h&) = GRID.TextMatrix(i&, h&)
            Next h&
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
