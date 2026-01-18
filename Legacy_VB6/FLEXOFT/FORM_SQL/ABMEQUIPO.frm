VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form ABMEQUIPO 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "A.B.M. Equpos"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11505
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   11505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command43 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   10920
      Picture         =   "ABMEQUIPO.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Refrescar"
      Top             =   20
      Width           =   405
   End
   Begin VB.CommandButton Command2 
      Height          =   975
      Left            =   10200
      Picture         =   "ABMEQUIPO.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   0
      Width           =   615
   End
   Begin VB.TextBox TEXT4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
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
      Index           =   1
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "Modelo"
      Top             =   300
      Width           =   3375
   End
   Begin VB.TextBox TEXT4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
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
      Index           =   2
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "Tipo de Equipo"
      Top             =   300
      Width           =   3375
   End
   Begin VB.TextBox TEXT4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
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
      Index           =   3
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "Marca"
      Top             =   300
      Width           =   3375
   End
   Begin VB.TextBox TXTMODELO 
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
      Left            =   6720
      MaxLength       =   32
      TabIndex        =   2
      Top             =   600
      Width           =   3375
   End
   Begin VB.TextBox TXTTIPOEQUIPO 
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
      Left            =   3360
      MaxLength       =   32
      TabIndex        =   1
      Top             =   600
      Width           =   3375
   End
   Begin VB.TextBox TXTMARCA 
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
      Left            =   0
      MaxLength       =   32
      TabIndex        =   0
      Top             =   600
      Width           =   3375
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8280
      Top             =   7320
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
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3720
      OleObjectBlob   =   "ABMEQUIPO.frx":0A2B
      Top             =   4080
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   3600
      OleObjectBlob   =   "ABMEQUIPO.frx":0C5F
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   4605
      Left            =   0
      TabIndex        =   8
      Top             =   1200
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   8123
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   ""
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
   Begin VB.Label LBLID 
      BorderStyle     =   1  'Fixed Single
      Height          =   225
      Left            =   0
      TabIndex        =   7
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Label5"
      Height          =   240
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mnuImprimir 
         Caption         =   "Exportar a Excel"
      End
   End
End
Attribute VB_Name = "ABMEQUIPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Sub BORRACOMBINACION(ID&)
   'VALIDAR SI EL MISMO EXISTE EN LA TABAL DE CLIENTES
   CONDI$ = "ID = " & ID&
   Call BORRAREGISTROS("EQUIPOAHP", CONDI$, REGAF&, "NOMESS")
End Sub

Private Sub BOTEXIT_Click()
  Unload Me
End Sub

Sub CARGAGRILLAEQUI(Optional MARCAX$, Optional TEQUIPOX$, Optional MODELOX$)
     MSF.Rows = 1
     'VA FILTRANDO DE ACUERDO A LOS DATOS QUE HAYA O NO EN LAS CAJAS DE TEXTO
     MARCAX$ = TRIM$(MARCAX$): TEQUIPOX$ = TRIM$(TEQUIPOX$): MODELOX$ = TRIM$(MODELOX$)
     Call PoneEnCeroNull("EQUIPOAHP", "STATUS", REGAF&, "NOMESS")
     
     SQLX$ = "SELECT * FROM EQUIPOAHP WITH(NOLOCK)"
     SQLX$ = SQLX$ + " WHERE ID > 0 AND STATUS >= 0"
     If TRIM$(MARCAX$) <> "" Then SQLX$ = SQLX$ + " AND MARCA LIKE '" & MARCAX$ & "%'"
     If TRIM$(TEQUIPOX$) <> "" Then SQLX$ = SQLX$ + " AND T_EQUIPO LIKE '" & TEQUIPOX$ & "%'"
     If TRIM$(MODELOX$) <> "" Then SQLX$ = SQLX$ + " AND MODELO LIKE '" & MODELOX$ & "%'"
     SQLX$ = SQLX$ + " ORDER BY MODELO"
     Set RS = READSET(SQLX$)
     With RS
        Do While Not .EOF
           j& = j& + 1
           MSF.Rows = j& + 1
           MSF.TextMatrix(j&, 1) = SAFETEXT$(!ID)
           MSF.TextMatrix(j&, 2) = SAFETEXT$(!MARCA)
           MSF.TextMatrix(j&, 3) = SAFETEXT$(!T_EQUIPO)
           MSF.TextMatrix(j&, 4) = SAFETEXT$(!MODELO)
          .MoveNext
        Loop
     End With
     RS.Close
     Set RS = Nothing
End Sub

Private Sub CMBFILTRO_Change()
   TXTBUSCAR_Change
End Sub

Private Sub Command1_Click()
   Command1.Enabled = False
   '
   

99 Command1.Enabled = True

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
Private Sub GRID_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = 2 Then
    '
    I& = GRID.MouseRow
    j& = GRID.MouseCol
    If I& = 0 Or I& > GRID.Rows Then Exit Sub
    '
    On Error Resume Next
    GRID.Row = I&
    GRID.COL = j&
    On Error Resume Next

    NPX& = TRIM$(GRID.TextMatrix(I&, 1))
    If NPX& < 1 Then
        Exit Sub
    End If
   End If
End Sub
Private Sub MSF_2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

End Sub

Private Sub msf2_Click()
    Static MODO  As Boolean
    If (msf2.MouseRow = 0) Then
          
        ' Ordena en forma ascendente
        If MODO Then
            msf2.COL = msf2.MouseCol
            msf2.Sort = 2
            MODO = False
        ' Ordena en forma descendente
        Else
            msf2.COL = msf2.MouseCol
            msf2.Sort = 1
            MODO = True
        End If
    End If

End Sub

Private Sub MSF2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = 2 Then
    '
    I& = msf2.MouseRow
    j& = msf2.MouseCol
    If I& = 0 Or I& > msf2.Rows Then Exit Sub
    '
    On Error Resume Next
    msf2.Row = I&
    msf2.COL = j&
    On Error Resume Next

    NPX& = TRIM$(msf2.TextMatrix(I&, 1))
    If NPX& < 1 Then
        Exit Sub
    End If
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

Private Sub Command2_Click()
   Command2.Enabled = False
'   Call BORRAREGISTROS("EQUIPOAHP", "ID > 0", REGAF&)
   MARCAX$ = TRIM$(UCase$(TXTMARCA))
   TIPOEQUI$ = TRIM$(UCase$(Me.TXTTIPOEQUIPO))
   MODELOX$ = TRIM$(UCase$(Me.TXTMODELO))
   If MARCAX$ = "" Then
     Call COMUNI("Falta Marca")
     GoTo 99
   End If
   '
   If TIPOEQUI$ = "" Then
     Call COMUNI("Falta Tipo de Equipo")
     GoTo 99
   End If

   If MODELOX$ = "" Then
     Call COMUNI("Falta Modelo")
     GoTo 99
   End If

   CONDI$ = "MARCA = '" & MARCAX$ & "'"
   CONDI$ = CONDI$ + " AND T_EQUIPO = '" & TIPOEQUI$ & "'"
   CONDI$ = CONDI$ + " AND MODELO = '" & MODELOX$ & "'"
   ID1& = XVALO(VALOBADA("EQUIPOAHP", "ID", CONDI$, "NOMESS"))
   '
   If ID1& < 1 Then
50    NUEVO_ID& = 1 + XVALO(VALOBADA("EQUIPOAHP", "MAX(ID)", " ID > 0 ", "NOMESS"))
      SQLX$ = "INSERT INTO EQUIPOAHP (MARCA, T_EQUIPO, MODELO,ID)"
      SQLX$ = SQLX$ + " VALUES ( "
      SQLX$ = SQLX$ + " '" & MARCAX$ & "' "
      SQLX$ = SQLX$ + ", '" & TIPOEQUI$ & "'"
      SQLX$ = SQLX$ + ", '" & MODELOX$ & "'"
      SQLX$ = SQLX$ + ", " & NUEVO_ID&
      SQLX$ = SQLX$ + ")"
      FLEXCONN.Execute (SQLX$)
   Else
60    SQLX$ = "UPDATE EQUIPOAHP SET MARCA = '" & MARCAX$ & "'"
      SQLX$ = SQLX$ + ", T_EQUIPO = '" & TIPOEQUI$ & "'"
      SQLX$ = SQLX$ + ", MODELO = '" & MODELOX$ & "'"
      SQLX$ = SQLX$ + " WHERE ID = " & ID1&
      
      FLEXCONN.Execute (SQLX$)
   End If
   '
   Call CARGAGRILLAEQUI
   Me.TXTMARCA = ""
   Me.TXTMODELO = ""
   Me.TXTTIPOEQUIPO = ""
   Me.LBLID = ""
   On Error Resume Next
   Me.TXTMARCA.SetFocus
   On Error GoTo 0
99 Command2.Enabled = True
   
End Sub

Private Sub Command43_Click()
   Me.TXTMARCA = ""
   Me.TXTMODELO = ""
   Me.TXTTIPOEQUIPO = ""
   Me.LBLID = ""
   Call CARGAGRILLAEQUI
End Sub

Private Sub Form_Load()
   Call CARGA_ENCABEZADO(MSF, "ID/F0;Marca/A40;Tipo de Equipo/A40;Modelo/A40", Me)
  
   Call CARGAGRILLAEQUI(MARCAX$, TEQUIPOX$, MODELOX$)
End Sub

Private Sub Label7_Click()

End Sub

Private Sub mnuImprimir_Click()
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
      Call GENEXCEL(AREXPO$, MSF)
   End If
   '
99
End Sub

Private Sub MSF_DblClick()
     REG& = XVALO(MSF.MouseRow)
     Me.LBLID = MSF.TextMatrix(REG&, 1)
     MARCAX$ = MSF.TextMatrix(REG&, 2)
     TIEQUIX$ = MSF.TextMatrix(REG&, 3)
     MODELOX$ = MSF.TextMatrix(REG&, 4)
     '
     Me.TXTMARCA = MARCAX$
     Me.TXTTIPOEQUIPO = TIEQUIX$
     Me.TXTMODELO = MODELOX$
End Sub

Private Sub MSF_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        REG& = MSF.MouseRow
        If REG& < 1 Then Exit Sub
        MSF.TopRow = REG&
        MSF.Row = REG&
        MSF.ColSel = 1
        On Error Resume Next
        MSF.SetFocus
        On Error GoTo 0

        IDX& = MSF.TextMatrix(REG&, 1)
        OPX% = COMALTER%("Desea Borrar El Item Seleccionado\\Cancelar\Eliminar")
        If OPX% < 2 Then Exit Sub
        If OPX% = 2 Then
           CONDI$ = "IDCOMBI = " & IDX&
           If CantRegistros&("EQUICLI", CONDI$, "NOMESS") > 0 Then
              CONDI$ = "ID = " & IDX&
              Call ACTUREGISTRO("EQUIPOAHP", "STATUS", CONDI$, (-1), REGAF&, "NOMESS")
           Else
              Call BORRACOMBINACION(IDX&)
           End If
           Call CARGAGRILLAEQUI
        End If
     End If
End Sub
Private Sub TXTMARCA_Change()
    Call CARGAGRILLAEQUI(Me.TXTMARCA, Me.TXTTIPOEQUIPO, Me.TXTMODELO)
End Sub


Private Sub TXTMARCA_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       On Error Resume Next
       Me.TXTTIPOEQUIPO.SetFocus
       On Error GoTo 0
    End If
End Sub


Private Sub TXTMODELO_Change()
    Call CARGAGRILLAEQUI(Me.TXTMARCA, Me.TXTTIPOEQUIPO, Me.TXTMODELO)
End Sub

Private Sub TXTMODELO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       On Error Resume Next
       Me.Command2.SetFocus
       On Error GoTo 0
    End If

End Sub


Private Sub TXTTIPOEQUIPO_Change()
    Call CARGAGRILLAEQUI(Me.TXTMARCA, Me.TXTTIPOEQUIPO, Me.TXTMODELO)

End Sub


Private Sub TXTTIPOEQUIPO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       On Error Resume Next
       Me.TXTMODELO.SetFocus
       On Error GoTo 0
    End If

End Sub


