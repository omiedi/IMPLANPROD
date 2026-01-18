VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form RepoComp 
   BackColor       =   &H00C0D0D0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cronograma de Entregas"
   ClientHeight    =   5985
   ClientLeft      =   990
   ClientTop       =   1095
   ClientWidth     =   18795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   18795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   12120
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox TXTECODESTINO 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2595
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   8175
   End
   Begin VB.CommandButton Command6 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2400
      TabIndex        =   6
      ToolTipText     =   "Despliega Lista de Selección"
      Top             =   120
      Width           =   160
   End
   Begin VB.TextBox TXTDESTINO 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   1170
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
      OleObjectBlob   =   "RepoComp.frx":0000
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   3240
      OleObjectBlob   =   "RepoComp.frx":0234
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin MSFlexGridLib.MSFlexGrid MSF0 
      Height          =   5235
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   18555
      _ExtentX        =   32729
      _ExtentY        =   9234
      _Version        =   393216
      Rows            =   26
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      WordWrap        =   -1  'True
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSF_2 
      Height          =   3675
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   13875
      _ExtentX        =   24474
      _ExtentY        =   6482
      _Version        =   393216
      Rows            =   26
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de Hoy"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10920
      TabIndex        =   11
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Destino / Cliente"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   40
      Width           =   855
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1320
      TabIndex        =   9
      Top             =   480
      Visible         =   0   'False
      Width           =   3135
      WordWrap        =   -1  'True
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
      Begin VB.Menu mnuExportaExcel 
         Caption         =   "Exportar a Excel"
      End
   End
End
Attribute VB_Name = "RepoComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Sub ACTUALIZA_FILA(MSF As msFlexGrid, REG&)
   For i& = 1 To MSF.Cols - 1
      MSF.TextMatrix(REG&, i&) = REPROGRID.TXTCAJA(i&)
   Next i&
   MSF.TopRow = REG&
   MSF.Row = REG&
   MSF.ColSel = 1
   On Error Resume Next
   MSF.SetFocus
   On Error GoTo 0
   
End Sub

Private Sub Form_Load()
     
     Campos$ = "Cliente/A16;Vto.Reprogramado/D12;Vencimiento/D12;N° Plano/A16;Descripción/A24;O-C N°/A16;Cantidad/F9.1;"
     Campos$ = Campos$ + "Observaciones/A32;Operación/A18;F.Entregado/D12;Operador/A12"
     Campos$ = Campos$ + ";Días-Vto./F9.1;Entrega Anticipada/f13.1;Retraso en Días/F13.1;Entrega Prod. a Calidad/D20;Motivo de Atraso-Otros/A24;"
     Campos$ = Campos$ + "Entrega Material/D12;Comentarios/A24;Operador/A16;Descripción/A20;Hs/A12;Fecha/D8;Est./F9.1;Realizado/F9.1;Diferencia/F9.1;N° PIN/F9;"
     Campos$ = Campos$ + "Controlador/A18;Comentarios/A24;Ord. Fabricación/A0"
     Call CARGA_ENCABEZADO(Me.MSF0, Campos$, Me)
     MSF0.Rows = 3

     MSF0.MergeRow(0) = True
     MSF0.MergeCells = flexMergeFree
     MSF0.FixedRows = 2
     'PASO EL ENCABEZADO A LA SEGUNDA FILA Y AGREGO OTRO TEXTO A LA PRIMERA PARA COMBINAR CELDAS DE LA FILA 0
     For i& = 1 To MSF0.Cols - 1
       MSF0.TextMatrix(1, i&) = MSF0.TextMatrix(0, i&)
       Select Case i&
           Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17
              MSF0.ColAlignment(i&) = flexAlignLeftTop
              MSF0.TextMatrix(0, i&) = "CRONOGRAMA DE ENTREGAS"
           Case 18
              MSF0.TextMatrix(0, i&) = "COMENTARIOS"
              MSF0.ColAlignment(i&) = flexAlignLeftTop
           Case 19, 20, 21
              MSF0.TextMatrix(0, i&) = "RETRABAJOS"
              MSF0.ColAlignment(i&) = flexAlignLeftTop
              Case 22, 23, 24, 25, 26, 27, 28
              MSF0.TextMatrix(0, i&) = "CONTROL DE CALIDAD"
              MSF0.ColAlignment(i&) = flexAlignLeftTop
           Case Else
             MSF0.TextMatrix(0, i&) = " "
       End Select
     Next i&
     Text1 = FECHATEX$(HOY(HO$))
     '

End Sub




Private Sub Command6_Click()
  '
  Call SELECHO("DESTFAB")
  COD$ = VALACT1$("DESTFAB")
  Me.TXTDESTINO = COD$
  '
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    MSF0.Height = Me.Height - 1500
    
    MSF0.Width = Me.Width - 235
    If MSF02.Visible = True Then
        MSF0.Height = Me.Height - 1800
    End If
    On Error GoTo 0

End Sub

Private Sub mnuExportaExcel_Click()
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
      Call GENEXCEL(AREXPO$, MSF0)
   End If
   '
99

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
      Dim objexcel As Object, OBJTEMP As Object
   Set objexcel = CreateObject("Excel.Application")
   objexcel.Workbooks.Add

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
          If RENG1& = 1 Then
   objexcel.Range("A1:AC1").Select
   Selection.Merge
   Call ColorearRangoCeldasExcel("A1:AC1", 43, objexcel, XLBOOK, XLSHEET, 1)
            If VALORX = VALORX_ANT And TRIM$(VALORX) <> "" Then
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
               GoTo 10
             End If
          End If
          '
          XLSHEET.Cells(RENG1&, COLUM&) = VALORX
          XLSHEET.Cells(RENG1&, COLUM&).EntireColumn.AutoFit
          XLSHEET.Cells(RENG1&, COLUM&).EntireRow.AutoFit
          
          If TipoDato% = 1 Then XLSHEET.Cells(RENG1&, COLUM&).NumberFormat = "dd/mm/yyyy"
          'If TIPODATO% = 2 Then XLSHEET.Cells(RENG1&, COLU1&).NumberFormat = "#,##0.00"
          If TipoDato% = 0 Then XLSHEET.Cells(RENG1&, COLUM&).NumberFormat = "@"
          TipoDato% = 0
          VALORX_ANT = VALORX
10     Next COLU1&
   Next i&
   '
   Close #n1%
   On Error GoTo 30
   XLSHEET.SaveAs NombreDeArchivo$
   On Error GoTo 0
   '
   XLAPP.QUIT      ' cierra EXCEL que se esta ejecutando minimizado
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
   XLAPP.QUIT ' Cierra Excel
   Set XLAPP = Nothing
   Set XLBOOK = Nothing
   Set XLSHEET = Nothing
   '
   Call MENSERR(24, "Imposible Crear Planilla Excel.")
   Exit Sub
   '
End Sub




Private Sub MSF0_DblClick()
     REG& = MSF0.MouseRow
     COL& = MSF0.MouseCol
     '
     ORIGINALWIDTH = 7185
     'AGREGO LOS CAJAS DE TEXTOS Y LOS LABEL CON SU ROTULO
     Dim FORMX$(), NOCAMPO$(), EDITABLE$(), MEMORIZA$()
     ReDim Preserve MEMORIZA$(1)
     ReDim Preserve NOCAMPO$(1), FORMX$(1)
     '
10   CANCA% = 0
     Campos$ = "Cliente/A16;Vto.Reprogramado/D12;Vencimiento/D12@;N° Plano/A16;Descripción/A24;O-C N°/A16@;Cantidad/F9.1;"
     Campos$ = Campos$ + "Observaciones/A32;Operación/A18;F.Entregado/D12@;Operador/A12"
     Campos$ = Campos$ + ";Dias-Vto./F9.1;Entrega Anticipada/f13.1;Retraso en Días/F13.1;Entrega Producción a Calidad/D20;Motivo Atraso-Otros/A24@;"
     Campos$ = Campos$ + "Entrega Materiales/D12@;Comentarios/A24;Operador/A16;Descripción/A20;Hs/A12;Fecha/D8;Est./F9.1;Realizado/F9.1;Didferencia/F9.1;N°PIN/F9;"
     Campos$ = Campos$ + "Controlador/A18;Comentarios/A24;O-Fabricación/A0"
     Call TEXTOLOTES$(Campos$, ";")
     
     If UBound(CADENATEX$) > 40 Then
        Call COMUNI("Demasiados Campos (Máximo 40)")
        Exit Sub
     End If
     
     For i& = 1 To UBound(CADENATEX$)
        TEXLEIDO$ = (CADENATEX$(i&))
        PPXX% = InStr(TEXLEIDO$, "/")
        If PPXX% < 1 Then PPXX% = 1 + Len(TEXLEIDO$)
        CANCA% = CANCA% + 1
        If CANCA% >= UBound(NOCAMPO$) Then
          ReDim Preserve NOCAMPO$(CANCA%), FORMX$(CANCA%), EDITABLE$(CANCA%)
        End If
        '
        NOCAMPO$(CANCA%) = Left$(TEXLEIDO$, PPXX% - 1)
        FORMX$(CANCA%) = Mid$(TEXLEIDO$, PPXX% + 1)
        EDITABLE(CANCA%) = Right(TEXLEIDO$, 1)
        Load REPROGRID.TXTLABEL(CANCA%)
        Load REPROGRID.TXTCAJA(CANCA%)
        REPROGRID.TXTLABEL(CANCA%) = NOCAMPO$(CANCA%)
        REPROGRID.TXTCAJA(CANCA%).TEXT = ""
        FORMAXTOX$ = UCase$(Left$(FORMX$(CANCA%), 1))
        If FORMAXTOX$ = "A" Or FORMAXTOX$ = "D" Then
           Align1% = 0
        ElseIf FORMAXTOX$ = "F" Or FORMAXTOX$ = "I" Or FORMAXTOX$ = "G" Then
           Align1% = 1
        Else
           Align1% = 2
        End If
        REPROGRID.TXTCAJA(CANCA%).Alignment = Align1%
        If CANCA% > 1 And CANCA% <= 20 Then
           REPROGRID.TXTCAJA(CANCA%).Top = REPROGRID.TXTCAJA(CANCA% - 1).Top + 400
           REPROGRID.TXTLABEL(CANCA%).Top = REPROGRID.TXTLABEL(CANCA% - 1).Top + 400
        ElseIf CANCA% > 20 Then
           If CANCA% = 21 Then
              REPROGRID.Width = (ORIGINALWIDTH * 2)
           End If
           REPROGRID.TXTCAJA(CANCA%).Top = REPROGRID.TXTCAJA(CANCA% - 20).Top
           REPROGRID.TXTLABEL(CANCA%).Top = REPROGRID.TXTLABEL(CANCA% - 20).Top
           REPROGRID.TXTLABEL(CANCA%).Left = REPROGRID.TXTLABEL(CANCA%).Width + REPROGRID.TXTCAJA(CANCA%).Width + 250
           REPROGRID.TXTCAJA(CANCA%).Left = (REPROGRID.TXTLABEL(CANCA%).Width * 2) + REPROGRID.TXTCAJA(CANCA%).Width + 300
        End If
        '
        If EDITABLE(CANCA%) = "@" Then
           REPROGRID.TXTCAJA(CANCA%).Locked = False
           REPROGRID.TXTCAJA(CANCA%).BackColor = &HFFFFFF
        End If
        '
        REPROGRID.TXTCAJA(CANCA%).Visible = True
        REPROGRID.TXTLABEL(CANCA%).Visible = True
     Next i&
     '
     '
     'AGREGO EL VALOR A LAS CAJAS DE TEXTO
     'SI VIENE REBOTADO DE VALIDACION MUESTRA LOS DATOS DE LAS CAJAS DE TEXTO (LEYENDO EL VECTOR) CASO CONTRARIO TOMA LA DE LA GRILLA.
     If UBound(MEMORIZA$) > 1 Then
         For i& = 1 To REPROGRID.TXTCAJA.Count - 1
           REPROGRID.TXTCAJA(i&) = MEMORIZA(i&)
         Next i&
         Erase MEMORIZA$
     Else
         For i& = 1 To MSF0.Cols - 1
           REPROGRID.TXTCAJA(i&) = MSF0.TextMatrix(REG&, i&)
         Next i&
     End If
     
     REPROGRID.Show 1
     If REPROGRID.GUARDAR = 1 Then
        Call GUARDAR_INFO(OK%)
        If OK% >= 0 Then
          Call ACTUALIZA_FILA(MSF0, REG&)
          Unload REPROGRID
        Else
          For i& = 1 To REPROGRID.TXTCAJA.Count - 1
            'GUARDO LOS DATOS EN UN VECTOR PARA MOSTRAR LA INFORMACION CORREGIDA EN EL FRM REPROGRID Y LO VUELVA A MOSTRAR EN EL CASO QUE OK% SEA NEGATIVO
            ReDim Preserve MEMORIZA(i&)
            MEMORIZA(i&) = REPROGRID.TXTCAJA(i&)
            Unload REPROGRID.TXTLABEL(i&)
            Unload REPROGRID.TXTCAJA(i&)
          Next i&
          GoTo 10
        End If
     End If
End Sub
Sub GUARDAR_INFO(OK%)
     OK% = -1
     Campos$ = "Cliente/A16;VTO.REPROGRAMADO/D12;VENCIMIENTO/D12@;N° PLANO/A16;DESCRIPCION/A24;O-C N°/A16@;CANTIDAD/F9.1;"
     Campos$ = Campos$ + "OBSERVACIONES/A32;OPERAC./A18;F.ENTREGADO/D12@;OPERADOR/A12"
     Campos$ = Campos$ + ";DIAS-VTO./F9.1;ENTR.ANT./f13.1;RETR EN DIAS/F13.1;F.ENTR.PROD. A CALIDAD/D20;MOTIVO DE ATRASO-OTROS/A24@;"
     Campos$ = Campos$ + "F.ENTR. MAT./D12@;COMENTARIOS/A24;OPERADOR/A16;DESCRIPCION/A20;HS/A12;FECHA/D8;EST./F9.1;REAL./F9.1;DIF./F9.1;N°PIN/F9;"
     Campos$ = Campos$ + "CONTROLADOR/A18;COMENTARIOS/A24"

     For i& = 0 To REPROGRID.TXTCAJA.UBound
       If REPROGRID.TXTCAJA(i&).Locked = False Then
          Select Case REPROGRID.TXTLABEL(i&)
          Case "VENCIMIENTO"
            VENCIMIENTO = REPROGRID.TXTCAJA(i&)
            If TRIM$(VENCIMIENTO) <> "" Then
                If IsDate(VENCIMIENTO) Then
                   VENCIMIENTO = CVDAT(CVINT(VENCIMIENTO))
                Else
                   Call COMUNI("Dato guardado como 'VENCIMIENTO', Formato no Válido")
                   Exit Sub
                End If
            End If
            'GRABAR FECHA DE ENTREGA DE LA OF.
          Case "O-C N°"
            OCOMPRA$ = REPROGRID.TXTCAJA(i&)
            'GRABAR EN EL CAMPO Id_Orfa DE LA OF
          Case "F.ENTREGADO"
            FLANZA = REPROGRID.TXTCAJA(i&)
            If TRIM$(FLANZA) <> "" Then
                If IsDate(FLANZA) Then
                   FLANZA = CVDAT(CVINT(FLANZA))
                Else
                   Call COMUNI("Dato guardado como 'F.ENTREGADO' Formato no Válido")
                   Exit Sub
                End If
            End If
            'GRABAR EN EL CAMPO FECHA DE LANZAMIENTO DE LA OF
          Case "MOTIVO DE ATRASO-OTROS"
            OBSERVACIOX$ = REPROGRID.TXTCAJA(i&)
            'GRABAR LA OBSERVACION EN LA DE OF.
          Case "F.ENTR. MAT."
            F_ENTREGAMAT$ = REPROGRID.TXTCAJA(i&)
            If TRIM$(F_ENTREGAMAT$) <> "" Then
                If IsDate(F_ENTREGAMAT$) Then
                   F_ENTREGAMATERIAL% = FECHANUM(F_ENTREGAMAT$)
                Else
                   Call COMUNI("Dato guardado como 'F.ENTR. MAT.', Formato no Válido")
                   Exit Sub
                End If
            End If
          
         Case "N°PIN"
             PIN$ = REPOGRID.TXTCAJA(i&)
              
               'GRABAR LA F.ENTREGA DEL MATERIAL EN LA DE CONTROL FINAL
          End Select
       End If

     Next i&
     '
     NORFA$ = REPROGRID.TXTCAJA(29)

     SQLX$ = "SELECT FechEnSol, Id_Orfa, FechLan, Observa, NumeroChasis, FechEntMat  FROM ORDEFABR "
     SQLX$ = SQLX$ + " WHERE IDSUBOR = '" & NORFA$ & "'"
     '
     Dim RS As ADODB.Recordset
     Set RS = New ADODB.Recordset
     RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
       With RS 'OrdeFabr
         If Not (.EOF And .BOF) Then
           If TRIM$(VENCIMIENTO) <> "" Then !FechEnSol = VENCIMIENTO
           !ID_ORFA = Left$(OCOMPRA$, 32)
            If TRIM$(FLANZA) <> "" Then !FechLan = FLANZA
           !Observa = OBSERVACIOX$
           !NumeroChasis = Left$(TRIM$(PIN$), 24)
           !FechEntMat = CVDAT(F_ENTREGAMATERIAL%)
           .Update
         End If
     End With
     RS.Close
     Set RS = Nothing
     OK% = 0
       
End Sub

Private Sub MSF0_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

  With MSF0
    ' verifica que el MsFlexgrid contiene filas
    If .MouseRow = 1 And .MouseCol >= 0 Then
        ' le asigna el dato con TextMatrix
        Select Case .MouseCol
           Case 1
           .ToolTipText = "Cliente Destino"
           Case 2
           .ToolTipText = "Fecha de Finalización de la 'Ultima Operación, Viene del Válculo de Carga de Máquinas"
           Case 3
           .ToolTipText = "Fecha de Entrega (Viene de la Orden de Fabricación)"
           Case 4
           .ToolTipText = "Nro de Plano (Viene del Maestro de Artículos)"
           Case 5
           .ToolTipText = "Segunda Descripción del Item (Viene del Maestro de Artículos)"
           Case 6
           .ToolTipText = "Orden de Compra (Dato Cargado en la Generación de Orden de Compra Casillero Referencia)"
           Case 7
           .ToolTipText = "La Cantidad Proyectada a Realizar en la Generación de la Orden de Fabricación"
           Case 8
           .ToolTipText = "Primera Descripción del Item (Viene del Maestro de Artículos)"
           Case 9
           .ToolTipText = "Descripción de la Operación (Si yá estan todas Informadas Pone Texto: 'Pendiente Aprobación Calidad'"
           Case 10
           .ToolTipText = "Fecha de Lanzamiento de la Orden de Fabricación"
           Case 11
           .ToolTipText = "Operario/s de Operación Posterior a la Última Informada (Viene de los informes de Fabricación)"
           Case 12
           .ToolTipText = "Dias de Vencimiento (Diferencia de días entre el Vencimiento Reprogramado y la Fecha de Hoy"
           Case 13
           .ToolTipText = "Cantida de Unidades Cerradas (la of) Anteriores a la Fecha de Entrega de la Orden de Fabricación ([VENCIMIENTO])"
           Case 14
           .ToolTipText = "Diferencia de Días Entre la Fecha de Hoy y la Fecha de Entrega([VENCIMIENTO REPROGRAMDO] Y HOY)"
           Case 15
           .ToolTipText = "Sí la última Operación esté Informada, Presenta la Fecha de dicho Informe de Producción"
           Case 16
           .ToolTipText = "Observaciones Dato Editable (Doble Click)"
           Case 17
           .ToolTipText = "Fecha de Entrega del Material (Provisión de Componentes para dicha Orden) Dato Editable (Doble Click)"
           '
           Case Else
           .ToolTipText = .TextMatrix(.MouseRow, .MouseCol)
        End Select
        
    Else
        ' No muestra ningún ToolTip
        .ToolTipText = vbNullString
          
    End If
  End With
    
End Sub

Private Sub TXTDESTINO_Change()
    Me.TXTECODESTINO = TRIM$(ECOARCH$("DESTFAB", Me.TXTDESTINO))

    Call CARGAMSF0
'    Call CARGAMSF1
'    Call CARGAMSF2
'    Call CARGAMSF3
'    Call CARGA_VENCIMIENTO_MSF3
'    Call CARGATOTALES
End Sub
Sub CARGAMSF0()

    COLOR1& = &H8000000F: COLOR2& = &HEEEEDD
    ColorActivo& = &HC0E0DD
    
    MSF0.Rows = 2
    J& = 1
    FECHADEHOY% = FECHANUM(Text1)
    Me.TXTECODESTINO = TRIM$(ECOARCH$("DESTFAB", Me.TXTDESTINO))
    SQLX$ = " SELECT M.DESCRIDOS, M.DESCRIPCION,M.NPLANO ,M.Codigo, O.IDSUBOR,M.FAMILIA, O.CANPROY, O.CANAPRO, O.FechEnSol, Id_Orfa,M.CODINT, O.DESTINO, O.Fechcierr, O.FechEntMat, O.NumeroChasis"
    SQLX$ = SQLX$ + " FROM MASTER M WITH(NOLOCK) INNER JOIN ORDEFABR O WITH(NOLOCK)"
    SQLX$ = SQLX$ + " ON M.CODINT = O.COD_INTE"
    SQLX$ = SQLX$ + " WHERE O.CANAPRO > 0.05"
    SQLX$ = SQLX$ + " AND  O.DESTINO ='" & TRIM$(Me.TXTDESTINO) & "'"
    Set RS = READSET(SQLX$)
    With RS
        Do While Not .EOF
              CI1% = XVALO(!CODINT)
              DESTINO$ = Me.TXTECODESTINO
              NORFA$ = SAFETEXT$(!idsubor)
              VTO_REPROGR% = VTO_REPROGRAMADO%(NORFA$)
              FECHPRO = FECHATEX$(VTO_REPROGR%)
              FE% = CVINT(!FechEnSol)
              PLANOX$ = SAFETEXT$(!NPLANO)
              OCOMPRAZ$ = SAFETEXT$(!ID_ORFA)
              DESCRIX$ = SAFETEXT$(!DESCRIDOS)
              '
              NOPER% = OPERACIONPENDIENTE%(NORFA$)
              If NOPER% > 0 Then
                    CONDI$ = " CodIConj = " & CI1% & " AND NumeOper =" & NOPER%
                    Descrioper$ = VALOBADA("SECOPER", "descoper", CONDI$, "NOMESS")
              Else
                    Descrioper$ = "Pendiente Aprobación Calidad"
              End If
              '
              CONDI$ = "idsubor = '" & NORFA$ & "'"
              FECHENTREGA_PROD% = CVINT(VALOBADA("ORDEFABR", "FECHLAN", CONDI$, "NOMESS"))
              '
              CANTI = XVALO(!CanApro)
              OBSERVAX$ = SAFETEXT$(!Descripcion)
              CONDI$ = " Cod_INTE = " & CI1% & " AND NumeOper =" & NOPER% & " AND IDSUBOR = '" & NORFA$ & "' ORDER BY FEHORLIB DESC"
              CADENAOPERARIOS$ = VALOBADA("INFOFAB", "TOP 1 LEGAOPS", CONDI$, "NOMESS")
              If TRIM$(CADENAOPERARIOS$) <> "" Then CADENAOPERARIOS$ = NAMECADENAOPERARIOS$(CADENAOPERARIOS$)
                
              J& = J& + 1
              MSF0.Rows = J& + 1
              MSF0.TextMatrix(J&, 1) = DESTINO$
              MSF0.TextMatrix(J&, 2) = FECHATEX$(VTO_REPROGR%)
              MSF0.TextMatrix(J&, 3) = FECHATEX$(FE%)
              MSF0.TextMatrix(J&, 4) = PLANOX$
              MSF0.TextMatrix(J&, 5) = DESCRIX$
              MSF0.TextMatrix(J&, 6) = OCOMPRAZ$
              MSF0.TextMatrix(J&, 7) = FORMATNUM$(CANTI, "F10.1")
    
              MSF0.TextMatrix(J&, 8) = OBSERVAX$
              MSF0.TextMatrix(J&, 9) = Descrioper$
              MSF0.TextMatrix(J&, 10) = FECHATEX$(FECHENTREGA_PROD%)
              MSF0.TextMatrix(J&, 11) = CADENAOPERARIOS$
                DiasVencidos = DiferenciaDiasEntreFechas(FE%, FECHADEHOY%)
                If DiasVencidos < 0 Then DiasVencidos = 0
              MSF0.TextMatrix(J&, 12) = DiasVencidos
              MSF0.TextMatrix(J&, 13) = CantiOfCerradaAntesFecha(NORFA$, FE%) 'CANTIDAD DE LA OF CERRADA  ANTES DE LA FECHA DE VENCIMIENTO
              
              UltimaFechaDeCierre% = CVINT(!Fechcierr)
              If UltimaFechaDeCierre% > 0 Then
                    DiasRetrasados% = DiferenciaDiasEntreFechas(FE%, UltimaFechaDeCierre%)
                    If DiasRetrasados% < 0 Then DiasRetrasados% = 0
                    MSF0.TextMatrix(J&, 14) = DiasRetrasados%
              Else
                    MSF0.TextMatrix(J&, 14) = ""
              End If
              MSF0.TextMatrix(J&, 15) = UltimaFechaUltimaOperacion%(NORFA$)
              
              MSF0.TextMatrix(J&, 16) = Label7
              MSF0.TextMatrix(J&, 17) = FECHATEX$(FechEntMate%)
              
              ' COLOREA LA FILA
              If ColorActivo& = COLOR1& Then
                    ColorActivo& = COLOR2&
              Else
                    ColorActivo& = COLOR1&
              End If
              RangoFila$ = CStr(MSF0.Rows - 1) & ":" & CStr(MSF0.Rows - 1)
              Call DisenioDeCelda(MSF0, RangoFila$, "1:28", "", 0, False, False, False, 0, vbBlack, ColorActivo&)
                        
              FilaInicioDeLaOF& = J&
             
              '*** INFORMES DE RETRABAJOS
              FilaRetrabajo& = FilaInicioDeLaOF&
              
              Dim RstRetrabajo As ADODB.Recordset
              ConsultaSql$ = "SELECT * FROM INFOFAB WITH(NOLOCK) WHERE IDSUBOR='" & NORFA$ & "' AND RETRABA=1"
              Set RstRetrabajo = READSET(ConsultaSql$)
              With RstRetrabajo
                Do While Not .EOF
                    
                    If FilaRetrabajo& > J& Then
                        J& = J& + 1
                        MSF0.Rows = J& + 1
                        RangoFila$ = CStr(MSF0.Rows - 1) & ":" & CStr(MSF0.Rows - 1)
                        Call DisenioDeCelda(MSF0, RangoFila$, "1:28", "", 0, False, False, False, 0, vbBlack, ColorActivo&)
                    End If
                    
                    MSF0.TextMatrix(FilaRetrabajo&, 19) = SAFETEXT$(!LegaOps)
                    MSF0.TextMatrix(FilaRetrabajo&, 20) = SAFETEXT$(!descoper)
                        TiempoRetrabajo# = XVALO(!TiemEjec) / 60
                    MSF0.TextMatrix(FilaRetrabajo&, 21) = FORMATNUM$(TiempoRetrabajo#, "F10.2")
              
                    FilaRetrabajo& = FilaRetrabajo& + 1
                    .MoveNext
                Loop
              End With
              On Error Resume Next
              RstRetrabajo.Close
              Set RstRetrabajo = Nothing
              
              '*** INFORMES DE CALIDAD
              FilaCalidad& = FilaInicioDeLaOF&
                            
              Dim RstCalidad As ADODB.Recordset
              ConsultaSql$ = "SELECT * FROM INFOCALIDAD WITH(NOLOCK) WHERE DOCUMENTOORIGEN='" & NORFA$ & "'"
              Set RstCalidad = READSET(ConsultaSql$)
              With RstCalidad
                Do While Not .EOF
                    
                    If FilaCalidad& > J& Then
                        J& = J& + 1
                        MSF0.Rows = J& + 1
                        RangoFila$ = CStr(MSF0.Rows - 1) & ":" & CStr(MSF0.Rows - 1)
                        Call DisenioDeCelda(MSF0, RangoFila$, "1:28", "", 0, False, False, False, 0, vbBlack, ColorActivo&)
                    End If
                    
                    MSF0.TextMatrix(FilaCalidad&, 22) = FECHATEX$(CVINT(!FechaInforme))
                    MSF0.TextMatrix(FilaCalidad&, 23) = ""
                    MSF0.TextMatrix(FilaCalidad&, 24) = TRIM$(FORMATNUM$(!CantidadAprobada, "F10.1"))
                    MSF0.TextMatrix(FilaCalidad&, 25) = ""
                    MSF0.TextMatrix(FilaCalidad&, 26) = SAFETEXT$(!NumeroDeControl)
                        NumUsuario% = !NumeroUsuario
                    MSF0.TextMatrix(FilaCalidad&, 27) = ECOARCH$("FLEXUSR", MKI$(NumUsuario%))
                    MSF0.TextMatrix(FilaCalidad&, 28) = SAFETEXT$(!Observaciones)
              
                    FilaCalidad& = FilaCalidad& + 1
                    .MoveNext
                Loop
              End With
              On Error Resume Next
              RstCalidad.Close
              Set RstCalidad = Nothing
              On Error GoTo 0
              
              .MoveNext
        Loop
    End With
    RS.Close
    Set RS = Nothing

End Sub
Function CantiOfCerradaAntesFecha(NORFA$, HASTAX%)
   CONDI$ = "DOPRILAR = '" & NORFA$ & "' AND CODIMOVI = 'OF' AND FECHMOV <  '" & FETEXCOR1$(HASTAX%) & "'"
   Cantidad = XVALO(VALOBADA("MOVISTO", "SUM(CANTINGRE)", CONDI$, "NOMESS"))
   CantiOfCerradaAntesFecha = Cantidad
End Function
Function NAMECADENAOPERARIOS$(CADENAOPERARIOS$)
            VUELTA% = 0
            For U& = 0 To 3
             OPER% = XVALO(Mid$(CADENAOPERARIOS$, INI% + 1, 2))
             If OPER% > 0 Then
               If VUELTA% = 0 Then
                 VUELTA% = 1
                 NAMECADENAOPERARIOS$ = Descrioper$ + ECOARCH$("PADRON", MKI$(OPER%)) + ","
               Else
                 NAMECADENAOPERARIOS$ = Descrioper$ + "," + ECOARCH$("PADRON", MKI$(OPER%))
               End If
             End If
             INI% = INI% + 2
            Next U&
End Function

Function VTO_REPROGRAMADO%(NORFA$)
       RFF$ = RECFILT$("OPERFAP", 1, NORFA$)
       For KU& = 1 To Len(RFF$) Step 4
         TYU& = CVS(Mid$(RFF$, KU&, 4))
         XX$ = REGLEIDO$("OPERFAP", TYU&)
         FEX2% = CVI(Mid$(XX$, 139, 2))
         If FEX2% > FEXLEIDO% Then FEXLEIDO% = FEX2
         Debug.Print FECHATEX$(FEX2%) & "      " & FECHATEX$(FEXLEIDO%)
       Next KU&
       Call CIERRARCH("OPERFAP")
       VTO_REPROGRAMADO% = FEXLEIDO

End Function

