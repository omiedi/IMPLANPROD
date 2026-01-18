VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form NovAve09 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   14415
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   14415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "ActulizarPedido de Reposición en  Base a Orden de Venta Generada en Central"
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
      Left            =   7520
      TabIndex        =   1
      Top             =   7440
      Width           =   6885
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
      Left            =   4560
      ScaleHeight     =   165
      ScaleWidth      =   1110
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame2 
      Caption         =   "Pedidos de Reposición Pendientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   14430
      Begin VB.CommandButton Command1 
         Caption         =   "Exp. a Excel"
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
         Left            =   7920
         TabIndex        =   11
         Top             =   3480
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Frame Frame1 
         Height          =   600
         Left            =   120
         TabIndex        =   6
         Top             =   3360
         Width           =   7455
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
            TabIndex        =   8
            Top             =   160
            Width           =   3135
         End
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
            TabIndex        =   7
            Top             =   160
            Width           =   2535
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
            TabIndex        =   10
            Top             =   160
            Width           =   705
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
            TabIndex        =   9
            Top             =   160
            Width           =   585
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GRID 
         Height          =   3165
         Left            =   15
         TabIndex        =   12
         ToolTipText     =   "Un Click Sobre Encabezado Para Ordenar/Doble Click Para Pasar a Grilla Inferior/Click Boton Derecho Imagen Digital"
         Top             =   240
         Width           =   14355
         _ExtentX        =   25321
         _ExtentY        =   5583
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
      Begin MSFlexGridLib.MSFlexGrid MSF_2 
         Height          =   3165
         Left            =   20
         TabIndex        =   13
         ToolTipText     =   "Un Click Sobre Encabezado Para Ordenar/Doble Click Para Pasar a Grilla Inferior/Click Boton Derecho Imagen Digital"
         Top             =   240
         Visible         =   0   'False
         Width           =   14350
         _ExtentX        =   25321
         _ExtentY        =   5583
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
   End
   Begin VB.Frame Frame3 
      Caption         =   "Detalle dePedidos de Reposición Seleccionada vs. Pedidos  Generados en Central"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   0
      TabIndex        =   14
      Top             =   4080
      Width           =   14415
      Begin MSFlexGridLib.MSFlexGrid MSFDET 
         Height          =   3045
         Left            =   20
         TabIndex        =   15
         ToolTipText     =   "Un Click Sobre Encabezado Para Ordenar/Doble Click Para Pasar a Grilla Inferior/Click Boton Derecho Imagen Digital"
         Top             =   240
         Width           =   14350
         _ExtentX        =   25321
         _ExtentY        =   5371
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
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Height          =   225
      Left            =   8520
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Label5"
      Height          =   240
      Left            =   0
      TabIndex        =   3
      Top             =   1000
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   0
      TabIndex        =   2
      Top             =   1000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Menu menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mnucopiar 
         Caption         =   "Copiar Celda Seleccionada"
         Visible         =   0   'False
      End
      Begin VB.Menu S 
         Caption         =   "_"
      End
   End
End
Attribute VB_Name = "NovAve09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FrmLeft, FrmTop, FrmWidth, FrmHeight As Long
Dim GridTop, GridLeft, GridWidth, GridHeight As Long
Dim FRAMETop, FRAMELeft, FRAMEWidth, FRAMEHeight As Long
Dim Com1Top, Com2Top, Com2Left As Long
Dim Com1LEFT As Long
Public APROBO%
Dim LecturaDatosOriginales As Byte
Dim GLOBALFRM_strContenidoCelda$
Private Sub archivo_Click()

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

Sub GENEXCEL(NOMBREDEARCHIVO$, GRID As MSFlexGrid)
   Dim XLAPP As Excel.Application
   Dim XLBOOK As Excel.Workbook
   Dim XLSHEET As Excel.Worksheet
   '
   Dim FECHA As Date
   Dim VALOR As Double
   Dim TEXTO As String
   '
   On Error GoTo 20
   Set XLAPP = New Excel.Application
   Set XLBOOK = XLAPP.Workbooks.Add
   Set XLSHEET = XLBOOK.Worksheets.Add
   On Error GoTo 0
   '
   TIPODATO% = 0: COLUM& = 0
   For I& = 1 To GRID.Rows
       RENG1& = I&: COLUM& = 0
       For COLU1& = 1 To GRID.Cols
          If GRID.ColWidth(COLU1& - 1) = 0 Then GoTo 10 ' OBVIA LAS COLUMNAS OCULTAS
          COLUM& = COLUM& + 1
          '
          VALORX = GRID.TextMatrix(I& - 1, COLU1 - 1)
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
   Next I&
   '
   Close #N1%
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

Private Sub Command2_Click()
     Command2.Enabled = False
     On Error Resume Next
     NUOC& = XVALO(MSFDET.TextMatrix(1, 1))
     If NUOC& < 1 Then
        Call COMUNI("No hay Una Orden de Compra Activa")
        On Error GoTo 0
        GoTo 99
     End If
     On Error GoTo 0
     
     NPEDX& = XVALO(MSFDET.TextMatrix(1, 2))
     'VALIDO SI HAY DIFERENCIAS
     HAYDIFERENCIA% = 0
     For I& = 1 To MSFDET.Rows - 1
       If XVALO(MSFDET.TextMatrix(I&, 7)) > 0 Then
          HAYDIFERENCIA% = 1
          Exit For
       End If
     Next I&
     
     If HAYDIFERENCIA% < 1 Then
        Call COMUNI("No Se Detectaron Difrencias")
     End If
     '
     'CAMBIO LA REVISION EN OCOMENCA
     '******************************
     CONDI$ = "NUME_OCOM = " & NUOC&
     Call ACTUREGISTRO("OCOMENCA", "OCOMREVIS", CONDI$, 1 + OCOMREVIS, REGAF&, "NOMESS")
     '
     'LEO LOS REGISTROS DE OCOMDETA Y GENERO LOS NUEVOS REGISTROS COMO EN LA MODFICACION MARCANDO COMO ANULADOS LOS ULTIMOS Y GENERANDO LA INFO NUEVA
     '***********************************************************************************************************************************************
     SQLX$ = "SELECT * FROM OCOMDETA  "
     SQLX$ = SQLX$ + " WHERE Nume_Ocom = " & NUOC&
     SQLX$ = SQLX$ + " AND STAT_OCOM < 9 "
     Dim RS As ADODB.Recordset
     Set RS = New ADODB.Recordset
     RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
     With RS
       Do While Not .EOF
          'GUARDA TODOS LOS DATOS DE LA FILA
          '*********************************
          CACAM& = RS.Fields.Count
          ReDim VARS(CACAM&)
          For U& = 1 To CACAM&
             VARS(U&) = RS.Fields(U& - 1).Value
          Next U&
          .AddNew
           For U& = 1 To CACAM&
              RS.Fields(U& - 1).Value = VARS(U&)
              If RS.Fields(U& - 1).Name = "Stat_Ocom" Then
                 !Stat_Ocom = 10 'DE ESTE MODO ME VAN A QUEDAR LOS REGISTROS AGREGADOS CON STAT_OCOM = 10
              End If
           Next U&
           .Update
          .MoveNext
       Loop
     End With
     RS.Close
     Set RS = Nothing
     '
     'CONSULTO LA ORDEN DE COMPRA POR CODIGO Y NRO DE ORDEN DE COMPRA
     '_ ACTUALIZO EL PRIMER REGISTRO SI HAY MAS LO BORRO (TENER EN CUENTA LA VERSION Y EL STATUS )
     'LA INFORMACION LA TOMO DE LA GRILLA.
     '************************************************************************************************
       Dim RS1 As ADODB.Recordset
       Set RS1 = New ADODB.Recordset

      For j& = 1 To MSFDET.Rows - 1
         NUOC& = XVALO(MSFDET.TextMatrix(j&, 1))
         NPEDX& = XVALO(MSFDET.TextMatrix(j&, 2))
         CI1% = CODINT(MSFDET.TextMatrix(j&, 3))
         CODIGO$ = MSFDET.TextMatrix(j&, 3)
         CANTIPEDIDO = XVALO(MSFDET.TextMatrix(j&, 6))
         SQLX$ = "SELECT * FROM OCOMDETA  "
         SQLX$ = SQLX$ + " WHERE Nume_Ocom = " & NUOC&
         SQLX$ = SQLX$ + " AND STAT_OCOM < 9 "
         SQLX$ = SQLX$ + " AND COD_INTE = " & CI1%
         VUELTA% = 0
         RS1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
         With RS1
          Do While Not .EOF
           If VUELTA% = 0 Then
            !Cant_Ocom = CANTIPEDIDO
            !UNIDS_OCOM = CANTIPEDIDO
            VUELTA% = 1
            .Update
           Else 'SI HAY REGISTROS CON MISMO CODIGO LOS BORRA POR QUE EN LA GRILLA PRESENTA LA INFO CON UN SUM EJEMPLO SI HAY UN ITEM EN DOS FILAS CON 4 C/U
            .Delete                               '_VA A PRESENTAR SOLO UNO CON 8 U.
           End If
           .MoveNext
          Loop
         End With
      Next j&
      RS1.Close
      Set RS1 = Nothing
      
     Call FORCOMP07.NOVEDADESAVELLANEDA(0, 0)
99   Command2.Enabled = True
     
End Sub
Private Sub Form_Load()
  Call CREAR_TABLA_DetaOVenta
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  APROBO% = 0
End Sub

Private Sub GRID_Click()
   Call EVENTOCLICK(GRID)
End Sub
Sub EVENTOCLICK(GRID As MSFlexGrid)
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
     
     REG& = GRID.MouseRow
     COL% = GRID.MouseCol
     If REG& < 1 Then Exit Sub
     NUOC& = XVALO(GRID.TextMatrix(REG&, 1))
     NPEDX& = XVALO(GRID.TextMatrix(REG&, 5))
     
     '
     SQLX$ = "Select SUM(Cant_Ocom) AS CANTICOMPRA, Descritem,Cod_Inte,Cod_Exte from OCOMDETA WITH(NOLOCK)"
     SQLX$ = SQLX$ + " WHERE NUME_OCOM = " & NUOC&
'     SQLX$ = SQLX$ + "  AND Nume_Pedi = " & NPEDX&
     SQLX$ = SQLX$ + "  AND stat_ocom < 9"
     SQLX$ = SQLX$ + "   GROUP BY Cod_Inte,DESCRITEM,Cod_Exte"
     Set RS = READSET(SQLX$)

    'HASTA ACA LLEGUE VER COMO TRAER LA INFO DE ESTAS DOS TABLAS. Y COMO ORDENAR 12/11/15
     With RS
       Do While Not .EOF
         j& = j& + 1
         MSFDET.Rows = j& + 1
         CI1% = XVALO(!Cod_Inte)
         MSFDET.TextMatrix(j&, 1) = NUOC&
         MSFDET.TextMatrix(j&, 2) = NPEDX&
         MSFDET.TextMatrix(j&, 3) = SAFETEXT$(!Cod_Exte)
         MSFDET.TextMatrix(j&, 4) = SAFETEXT$(!DESCRITEM)
         MSFDET.TextMatrix(j&, 5) = SAFETEXT$(!CANTICOMPRA)
           CONDI$ = "NRO_OCOM = " & NUOC&
           CONDI$ = CONDI$ + " AND nroped = " & NPEDX&
           CONDI$ = CONDI$ + " AND CODINT = " & CI1%
           CANTIPEDIDO = XVALO(VALOBADA("DETAOVENTA", "sum(CANPED)", CONDI$, "NOMESS"))
         MSFDET.TextMatrix(j&, 6) = CANTIPEDIDO
         MSFDET.TextMatrix(j&, 7) = XVALO(MSFDET.TextMatrix(j&, 5)) - CANTIPEDIDO
        .MoveNext
       Loop
     End With

End Sub


Private Sub GRID_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = 2 Then
     REG& = GRID.MouseRow
     COL% = GRID.MouseCol
     Call MENU_CELDA(GRID, REG&, COL%)
  End If

End Sub

Private Sub MSF_2_Click()
    Call EVENTOCLICK(MSF_2)
End Sub

Sub INVICTRL()
  '
  mnucopiar.Visible = Not (mnucopiar.Visible)
  '
End Sub

Private Sub mnucopiar_Click()
       Clipboard.Clear
       Clipboard.SetText GLOBALFRM_strContenidoCelda$
End Sub

Private Sub MSFDET_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = 2 Then
     REG& = MSFDET.MouseRow
     COL% = MSFDET.MouseCol
     Call MENU_CELDA(MSFDET, REG&, COL%)
  End If
End Sub

Sub MENU_CELDA(MSF As MSFlexGrid, REG&, COL%)
        If MSF.Rows <= 1 Then Exit Sub
        Call COLOREAR_GRILLA_SOLACELDA(MSF, vbYellow, REG&, Int(COL%))   ' LA PINTA DE AMARILLA
        Call INVICTRL
        GLOBALFRM_strContenidoCelda$ = TRIM$(MSF.TextMatrix(REG&, COL%))  'CELDA SELECCIONADA
        PopupMenu menu
        Call INVICTRL
        Call COLOREAR_GRILLA_SOLACELDA(MSF, vbWhite, REG&, COL%)  ' LA PINTA DE AMARILLA
End Sub

'Private Sub MSF2_Click()
'    Static MODO  As Boolean
'    If (MSF2.MouseRow = 0) Then
'
'        ' Ordena en forma ascendente
'        If MODO Then
'            MSF2.COL = MSF2.MouseCol
'            MSF2.Sort = 2
'            MODO = False
'        ' Ordena en forma descendente
'        Else
'            MSF2.COL = MSF2.MouseCol
'            MSF2.Sort = 1
'            MODO = True
'        End If
'    End If
'
'End Sub



Private Sub TXTBUSCAR_Change()
    '
    If TRIM$(TXTBUSCAR) = "" Then
      MSF_2.Visible = False
      GRID.Visible = True
      Exit Sub
    End If
    '
    MSF_2.Rows = 2
    For I& = 1 To MSF_2.Cols - 1
       MSF_2.TextMatrix(1, I&) = ""
    Next I&
    REG& = 0
    MSF_2.Visible = True
    MSF_2.ZOrder 0
    MSF_2.Cols = GRID.Cols
    MSF_2.Redraw = False
    GRID.Visible = False
    '
    'RECORRO EL MSF QUE TIENE TODO E IMPRIMO EN msf  LOS QUE MACHEAN
    For I& = 1 To GRID.Rows - 1
      For j& = 1 To GRID.Cols - 1
        If (InStr(UCase$(GRID.TextMatrix(I&, j&)), UCase$(TXTBUSCAR))) Then
          If (TRIM$(CMBFILTRO.TEXT)) = "" Or TRIM$(UCase$(CMBFILTRO)) = TRIM$(UCase$(GRID.TextMatrix(0, j&))) Then
            REG& = REG& + 1
            MSF_2.Rows = REG& + 1
            For H& = 1 To MSF_2.Cols - 1
              MSF_2.TextMatrix(REG&, H&) = GRID.TextMatrix(I&, H&)
            Next H&
          End If
          If (TRIM$(CMBFILTRO.TEXT)) = "" Then Exit For
        End If
      Next j&
    Next I&
    '
    MSF_2.Redraw = True
    
    '
End Sub

Private Sub TXTBUSCAR_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      TXTBUSCAR.TEXT = ""
   End If
End Sub
