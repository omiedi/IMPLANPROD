VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form VERRECIBOS1 
   Caption         =   "Visualizar Recibos"
   ClientHeight    =   6825
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   9945
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4680
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exp.Excel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   7
      ToolTipText     =   "Exportar a Excel"
      Top             =   6120
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "VER TODOS "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7200
      TabIndex        =   6
      Top             =   6120
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox TEXT2 
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
      Left            =   4725
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   105
      Width           =   5130
   End
   Begin VB.TextBox Text1 
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
      Left            =   1785
      TabIndex        =   3
      Top             =   105
      Width           =   2595
   End
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
      Height          =   315
      Left            =   4410
      TabIndex        =   2
      Top             =   105
      Width           =   255
   End
   Begin VB.TextBox Text17 
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
      Index           =   6
      Left            =   105
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "Legajo:"
      Top             =   105
      Width           =   1605
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   5445
      Left            =   105
      TabIndex        =   0
      ToolTipText     =   "Doble click para Visualizar"
      Top             =   525
      Width           =   9750
      _ExtentX        =   17198
      _ExtentY        =   9604
      _Version        =   393216
      Cols            =   4
      BackColor       =   14737632
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      FormatString    =   "|>CAMPO1|>CAMPO2|>CAMPO3"
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "VERRECIBOS1.frx":0000
      Top             =   0
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "VERRECIBOS1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Command1.Enabled = False
    '
    Call ABRECONEXION
    '
    Call CARGA_INDICE_ALUMNOS
    '
    If RESPUESTA_BUSQUEDA$ = "" Then GoTo 99
    '
    Text1 = RESPUESTA_BUSQUEDA$
    
99  Command1.Enabled = True

End Sub

Private Sub Command2_Click()
   Text1 = ""
   Call CARGA_RECIBOS
End Sub

 Private Sub Command3_Click()
   Command3.Enabled = False
   
   'valido que haya informacion en el list
   If MSF.Rows <= 1 Then Exit Sub
   '
   'presenta cuadro de dialogo
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
   ARCHIEX$ = UCase$(CommonDialog1.FileName)
   
   'si selecciona archivo genera excel
   If TRIM$(ARCHIEX$) <> "" Then
      Call GENERAXCEL(VERRECIBOS.MSF, ARCHIEX$)
   End If
   
99 Command3.Enabled = True

End Sub


Private Sub Form_Load()
    Call APLICAR_SKIN(Me, App.Path & "\SKINS\SteelBlue.skn")
    Call ENCABEZA_GRILLA
'
End Sub

Private Sub MSF_DblClick()
   LGJO& = XVALO(MSF.TextMatrix(MSF.Row, 10))
   'ASIGNO POSICION SELECCIONADA A VARIABLES
   i& = MSF.MouseRow
   j& = MSF.MouseCol
   '
   If i& < 1 Then Exit Sub
   'VALIDO
   If i& = 0 Or i& > MSF.Rows Then Exit Sub
   If j& > MSF.Cols Then Exit Sub
   '
   'RESTAURO AL COLOR ORIGINAL LA COLUMNA 0
   MSF.COL = 0
   For H& = 0 To MSF.Rows - 1
     MSF.Row = H&
     MSF.CellBackColor = &H8000000F
   Next H&
   '
   'PINTO DE AMARILLO LA SELECCIONADA
   MSF.Row = i&
   MSF.CellBackColor = &H80FFFF   ' LA PINTA DE AMARILLA
   '
   If LGJO& < 1 Then Exit Sub
   '
   DOCU$ = TRIM$(MSF.TextMatrix(MSF.Row, 2))
   VARMOV$ = TRIM$(MSF.TextMatrix(MSF.Row, 5))
   NURECI& = XVALO(Mid$(DOCU$, 5))
   '
   If VARMOV$ = "COBMATRI" Then
     Call PRIMATRICULA(LGJO&, NURECI&, DOCU$, "COBMATRI")
   ElseIf VARMOV$ = "RCOB" Then
     Call PRIRECIBO(LGJO&, NURECI&, DOCU$, "RCOB")
   End If
   '
End Sub

Private Sub Text1_Change()
'     Call LIMPIAR
     Call CARGA_RECIBOS
     TEXT2 = VALOADMIN("ALUMNOS", "NOMBRE_APELLIDO", "LEGAJO = " & XVALO(Text1), "NOMESS")


End Sub
Sub CARGA_RECIBOS()
   '
   MSF.Rows = 1
   MSF.Cols = 11
   '
   LEGAJO& = XVALO(Text1)
   
   CONDI$ = " (VARIMOVI = 'RCOB' OR VARIMOVI = 'COBMATRI')"
   CONDI$ = CONDI$ + "  AND TIPOMOVIM = 'RECIBO COBRANZA'"
   If LEGAJO& > 0 Then CONDI$ = CONDI$ + " AND LEGAJO = " & LEGAJO& & ""
   '
   If CantRegistros&("CAJABANC_COL", CONDI$, "NOMESS") < 1 Then Exit Sub
   '
   MSF.Redraw = False
   '
   SQLX$ = "SELECT * FROM CAJABANC_COL "
   SQLX$ = SQLX$ + " WHERE " + CONDI$
   SQLX$ = SQLX$ + " order by fechemisi desc"
   Set RECITMP = BASSSEX.Execute(SQLX$)
'   Set RECITMP = BaDaABRE_ADMIN_ALQ.OpenRecordset(SQLX$, dbOpenSnapshot)
   With RECITMP
      Do While Not .EOF
           FEMI% = CVINT(!FECHEMISI)
           DOCUM$ = SAFETEXT$(!CODICOM_LAR)
           MEX% = XVALO(!MES_CUOTA)
           ANOX% = XVALO(!ANO_CUOTA)
           CONCEP = SAFETEXT$(!VARIMOVI)
           CUO_ORIGINALX# = XVALO(!CUOTA_ORIGINAL)
           BONIFICACIONX# = XVALO(!BONIFICACION)
           RECARGOX# = XVALO(!RECARGO)
           SUPAGO# = XVALO(!HABER)
           LGJO& = XVALO(!LEGAJO)
          '
           FILA& = FILA& + 1
           MSF.Rows = FILA& + 1
           MSF.TextMatrix(FILA&, 1) = FECHATEX$(FEMI%)
           MSF.TextMatrix(FILA&, 2) = DOCUM$
           MESXX$ = "00" & TRIM$(Str(MEX%))
           MESXX$ = Right(MESXX$, 2)
           MSF.TextMatrix(FILA&, 3) = MESXX$
           MSF.TextMatrix(FILA&, 4) = SAFETEXT$(ANOX%)
           MSF.TextMatrix(FILA&, 5) = CONCEP
           MSF.TextMatrix(FILA&, 6) = TRIM$(FORMATNUM$(CUO_ORIGINALX#, "F12.2"))
           MSF.TextMatrix(FILA&, 7) = TRIM$(FORMATNUM$(BONIFICACIONX#, "F12.2"))
           MSF.TextMatrix(FILA&, 8) = TRIM$(FORMATNUM$(RECARGOX#, "F12.2"))
           MSF.TextMatrix(FILA&, 9) = TRIM$(FORMATNUM$(SUPAGO#, "F12.2"))
           MSF.TextMatrix(FILA&, 10) = TRIM$(Str$(LGJO&))
           .MoveNext
      Loop
    End With
    RECITMP.Close
    Set RECITMP = Nothing
    '
    Call ENCABEZA_GRILLA
    MSF.Redraw = True



End Sub
Sub ENCABEZA_GRILLA()
    MSF.Cols = 11

    NUCOLU% = 1: TEPRUEBA = "10": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 1: MSF.TextMatrix(0, NUCOLU% - 1) = ""
    NUCOLU% = 2: TEPRUEBA = "10/10/20120": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 1: MSF.TextMatrix(0, NUCOLU% - 1) = "Fecha Emis."
    NUCOLU% = 3: TEPRUEBA = "12345678900000000000": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 1: MSF.TextMatrix(0, NUCOLU% - 1) = "Documento"
    NUCOLU% = 4: TEPRUEBA = "ABCD": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 1: MSF.TextMatrix(0, NUCOLU% - 1) = "Mes"
    NUCOLU% = 5: TEPRUEBA = "CHAPA": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 1: MSF.TextMatrix(0, NUCOLU% - 1) = "Año"
    NUCOLU% = 6: TEPRUEBA = "123456789012": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 1: MSF.TextMatrix(0, NUCOLU% - 1) = "Concepto"
    NUCOLU% = 7: TEPRUEBA = "123456789012": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 6: MSF.TextMatrix(0, NUCOLU% - 1) = "Imp.Original"
    NUCOLU% = 8: TEPRUEBA = "123456789012": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 6: MSF.TextMatrix(0, NUCOLU% - 1) = "Bonificación"
    NUCOLU% = 9: TEPRUEBA = "123456789012": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 6: MSF.TextMatrix(0, NUCOLU% - 1) = "Recargo"
    NUCOLU% = 10: TEPRUEBA = "123456789012": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 6: MSF.TextMatrix(0, NUCOLU% - 1) = "Su Pago"
    NUCOLU% = 11: TEPRUEBA = "": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 6: MSF.TextMatrix(0, NUCOLU% - 1) = ""

End Sub

