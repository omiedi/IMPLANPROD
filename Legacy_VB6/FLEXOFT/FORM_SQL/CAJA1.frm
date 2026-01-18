VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Begin VB.Form CAJA1 
   Caption         =   "Movimientos de Caja"
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   12660
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Actualizar Inforamación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   22
      Top             =   840
      Width           =   3255
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6120
      Top             =   3960
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
      Left            =   7920
      TabIndex        =   21
      ToolTipText     =   "Exportar a Excel"
      Top             =   7920
      Width           =   2175
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   6120
      OleObjectBlob   =   "CAJA1.frx":0000
      Top             =   720
   End
   Begin VB.CommandButton Command25 
      Caption         =   "Imprimir"
      Height          =   495
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7920
      Width           =   2415
   End
   Begin VB.TextBox Text20 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   3000
      MaxLength       =   50
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   7480
      Width           =   4815
   End
   Begin VB.TextBox Text20 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Index           =   0
      Left            =   3000
      MaxLength       =   50
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   6840
      Width           =   4815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Grabar Retiro"
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
      Index           =   0
      Left            =   7920
      TabIndex        =   19
      Top             =   6840
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   1680
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   7480
      Width           =   1215
   End
   Begin VB.TextBox Text16 
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
      Index           =   0
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Text            =   "Retiro de Caja:"
      Top             =   6840
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Index           =   0
      Left            =   1680
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   6840
      Width           =   1215
   End
   Begin VB.TextBox Text15 
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
      Index           =   3
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Text            =   "Ingreso de Caja:"
      Top             =   7480
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Grabar Ingreso Manual"
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
      Index           =   1
      Left            =   7920
      TabIndex        =   14
      Top             =   7480
      Width           =   2175
   End
   Begin VB.TextBox Text15 
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
      Index           =   8
      Left            =   10200
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Text            =   "Salidas:"
      Top             =   7025
      Width           =   1095
   End
   Begin VB.TextBox Text16 
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
      Index           =   2
      Left            =   10200
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Text            =   "Resultado:"
      Top             =   7440
      Width           =   1095
   End
   Begin VB.TextBox Text15 
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
      Index           =   0
      Left            =   11520
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   6600
      Width           =   1095
   End
   Begin VB.TextBox Text16 
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
      Index           =   1
      Left            =   10200
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "Ingreso:"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.TextBox Text15 
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
      Index           =   1
      Left            =   11520
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   7025
      Width           =   1095
   End
   Begin VB.TextBox Text15 
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
      Index           =   2
      Left            =   11520
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   7440
      Width           =   1095
   End
   Begin VB.TextBox Text17 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FBCA9D&
      Height          =   600
      Index           =   3
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "MOVIMIENTOS DE CAJA"
      Top             =   0
      Width           =   12675
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   1920
      TabIndex        =   2
      Top             =   620
      Width           =   10695
      Begin VB.TextBox Text16 
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
         Height          =   375
         Index           =   3
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Text            =   "Registros Generados"
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox TXTCANTIFILAS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker DESDE 
         Height          =   375
         Left            =   7080
         TabIndex        =   3
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CalendarBackColor=   14737632
         CalendarTitleBackColor=   14737632
         Format          =   130482177
         CurrentDate     =   40875
      End
      Begin MSComCtl2.DTPicker HASTA 
         Height          =   375
         Left            =   9000
         TabIndex        =   4
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CalendarBackColor=   14737632
         CalendarTitleBackColor=   14737632
         Format          =   130482177
         CurrentDate     =   40875
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   5085
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   12585
      _ExtentX        =   22199
      _ExtentY        =   8969
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColor       =   14737632
      BackColorFixed  =   16501405
      ForeColorSel    =   16501405
      BackColorBkg    =   16777215
      FocusRect       =   2
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "CAJA1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click(Index As Integer)
   Command1(Index).Enabled = False
   If TRIM$(Text20(Index)) = "" Then
        MsgBox "FALTA INGRESAR REFERENCIA"
        GoTo 99
   End If

   Select Case (Index)
     Case 0
        If XVALO(Text2(Index)) > 0 Then Call GRABA_MOVI_CAJA(XVALO(Text2(0)), 0)
     Case 1
        If XVALO(Text2(Index)) > 0 Then Call GRABA_MOVI_CAJA(0, XVALO(Text2(1)))
   End Select
   MsgBox "HECHO"
'   Call DESDE_Click
'   Call HASTA_Click
   Call Command2_Click
   Text2(0) = ""
   Text2(1) = ""
   Text20(0) = ""
   Text20(1) = ""
99 Command1(Index).Enabled = True

End Sub
Sub GRABA_MOVI_CAJA(Sali#, INGR#)
    
    SALIDA# = XVALO(Sali#)
    INGRESO# = XVALO(INGR#)
    
    If SALIDA# > 0 Then
       TiMovi$ = "SALIDA DE CAJA"
       VARIMOVIX$ = "SALIDA DE CAJA"
       REFE$ = TRIM$(Text20(0))
       NUMEDOCU& = 1 + MAXNUMCAMPO&("CAJABANC_COL", "NUMEDOC", "VARIMOVI = 'SALIDA DE CAJA' AND TIPOMOVIM = 'SALIDA DE CAJA'")
       CUECONTA% = 3
    ElseIf INGRESO# > 0 Then
       TiMovi$ = "INGRESO MANUAL"
       VARIMOVIX$ = "INGRESO MANUAL"
       REFE$ = TRIM$(Text20(1))
       NUMEDOCU& = 1 + MAXNUMCAMPO&("CAJABANC_COL", "NUMEDOC", "VARIMOVI = 'INGRESO MANUAL' AND TIPOMOVIM = 'INGRESO MANUAL'")

       CUECONTA% = 2
    End If
    
    SQLX$ = "SELECT * FROM CAJABANC_COL"
    SQLX$ = SQLX$ + " WHERE LEGAJO < 1"
      
    Set AUX = New ADODB.Recordset
    AUX.Open (SQLX$), BASSSEX, adOpenKeyset, adLockPessimistic, adCmdText
    With AUX
30        .AddNew
          !LEGAJO = 0
          !MES_CUOTA = 0
          !ANO_CUOTA = 0
          !TipoMovim = TiMovi$
          !REFERENCIA = Left$(REFE$, 64)
          !CUENTA_CONT = CUECONTA%
          !VARIMOVI = VARIMOVIX$
          !CODICOM_LAR = COMLAR$
          
          !NUMEDOC = NUMEDOCU&
          !HABER = SALIDA#
          !DEBE = INGRESO#
          !NUME_USUARIO = NUMEUSER%
          !FECHEMISI = CVDAT(HOY(HO$))
          !BONIFICACION = 0
          !RECARGO = 0
          !CUOTA_ORIGINAL = 0
          !ENSE_PROGRAMADA = 0
          !INGLES = 0
          !EMERGENCIA = 0
          !COMPUTACION = 0
          !DOCU_APLI = ""
          !FeReMovi = Now
          .Update
      End With
      AUX.Close
      Set AUX = Nothing
      
''      'GRABO NOTA DE ALUMNO
''      If Sali# > 0 Then
''        TX$ = TIMOVI$ & vbCrLf
''        TX$ = TX$ + "Referencia: " & REFE$ & vbCrLf
''        TX$ = TX$ + "Importe: $ " & TRIM$(FORMATNUM$(SALIDA#, "F15.2"))
''      ElseIf INGRESO# > 0 Then
''        TX$ = TIMOVI$ & vbCrLf
''        TX$ = TX$ + "Referencia: " & REFE$ & vbCrLf
''        TX$ = TX$ + "Importe: $ " & TRIM$(FORMATNUM$(INGRESO#, "F15.2"))
''      End If
''
''      Call ANOTAX_ALUMNO(LGJO&, TX$)

   
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   Call CARGA_GRILLA_CAJA
   Command2.Enabled = True
End Sub

Private Sub Command25_Click()
   Command25.Enabled = False
   'valido
   If MSF.Rows <= 1 Then
      MsgBox "GRILLA DE IMPRESIÓN VACIA"
      GoTo 99
   End If
   '
   Set GRILLA_IMPRESION = MSF
   '
   PREIMPRE.txtEncabezado2.TEXT = ""

   '
   
   ENCABEZADO1$ = "Reporte de Movimientos de Caja"
   
   PREIMPRE.txtEncabezado1.TEXT = ENCABEZADO1$
   PREIMPRE.txtPiePagina = "Fin de Reporte"
   PREIMPRE.Combo1.TEXT = "Vertical"
   PREIMPRE.Show 1
   '
99 Command25.Enabled = True

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
      Call GENERAXCEL(CAJA.MSF, ARCHIEX$)
   End If
   
99 Command3.Enabled = True

End Sub


Private Sub Form_Load()
    Call APLICAR_SKIN(Me, App.Path & "\SKINS\SteelBlue.skn")
    Call ENCABEZA_GRILLA
    DESDE.Value = Date
    HASTA.Value = Date
    Call CARGA_GRILLA_CAJA
    Call CARGAVECTOR_MES
    '
End Sub


Sub CARGA_GRILLA_CAJA()
   '
   MSF.Rows = 1
   MSF.Cols = 11
   Call LIMPIASUBTOTALES
   '
   Screen.MousePointer = 11
   DES% = CVINT(DESDE.Value)
   HAS% = CVINT(HASTA.Value)
   DESDEX = FETEXCOR1$(DES%)
   HASTAX = FETEXCOR1$(HAS%)
   FILA& = 0
   'CARGA EL ENCABEZADO DEL TRANSPORTE
   condi0$ = "(CUENTA_CONT = 2 OR CUENTA_CONT = 3)"
   condi0$ = condi0$ + " AND FECHEMISI < '" & DESDEX & "'"
   If CantRegistros&("CAJABANC_COL", condi0$, "NOMESS") < 1 Then GoTo 10
   
   SQLX$ = "SELECT SUM(DEBE) AS DEBEX, SUM(HABER)AS HABEX"
   SQLX$ = SQLX$ + " FROM CAJABANC_COL WITH(NOLOCK) "
   SQLX$ = SQLX$ + " WHERE " + condi0$
   SQLX$ = SQLX$ + " group by CUENTA_CONT"
   'Set RECITMP0 = BaDaABRE_ADMIN_ALQ.OpenRecordset(SQLX$, 8)

   Set RECITMP0 = BASSSEX.Execute(SQLX$)
   With RECITMP0
      Do While Not .EOF
           SUPAGO# = XVALO(!DEBEX)
           RETIRO# = XVALO(!HABEX)
           SUINGREACUM# = SUINGREACUM# + SUPAGO#
           RETIROACUM# = RETIROACUM# + RETIRO#
           '
           FILA& = FILA& + 1
           TXTCANTIFILAS = FILA&: TXTCANTIFILAS.Refresh
           MSF.Rows = FILA& + 1
           MSF.TextMatrix(FILA&, 1) = "-------------------"
           MSF.TextMatrix(FILA&, 2) = "-------------------"
           MSF.TextMatrix(FILA&, 3) = "-------------------"
           MSF.TextMatrix(FILA&, 4) = "-------------------"
           MSF.TextMatrix(FILA&, 5) = "Transporte................"
           MSF.TextMatrix(FILA&, 6) = "-------------------"
           MSF.TextMatrix(FILA&, 7) = TRIM$(FORMATNUM$(SUPAGO#, "F12.2"))
           MSF.TextMatrix(FILA&, 8) = TRIM$(FORMATNUM$(RETIRO#, "F12.2"))
           
           .MoveNext
      Loop
    End With
    RECITMP0.Close
    Set RECITMP0 = Nothing
    FILA& = FILA& + 1
    '
    'BLANQEO POR QUE PIDIERON QUE SOLO ACUMULE EL TOTAL DEL RANGO SELECCIONADO.
    SUINGREACUM# = 0
    RETIROACUM# = 0

10 CONDI$ = "(CUENTA_CONT = 2 OR CUENTA_CONT = 3) "
   CONDI$ = CONDI$ + " AND FECHEMISI >= '" & DESDEX & "' AND FECHEMISI <= '" & HASTAX & "'"
   '
   If CantRegistros&("CAJABANC_COL", CONDI$, "NOMESS") < 1 Then GoTo 90
   '
   MSF.Redraw = False
   '
   SQLX$ = "SELECT DEBE, HABER,FECHEMISI,REFERENCIA, CUOTA_ORIGINAL, A.LEGAJO, CODICOM_LAR, A.NOMBRE_APELLIDO   FROM CAJABANC_COL C WITH(NOLOCK)"
   SQLX$ = SQLX$ + " LEFT JOIN ALUMNOS A WITH(NOLOCK) ON C.LEGAJO = A.LEGAJO"
   SQLX$ = SQLX$ + " WHERE " + CONDI$
   SQLX$ = SQLX$ + " ORDER BY FECHEMISI, LEGAJO"
   Set RECITMP = BASSSEX.Execute(SQLX$)
   With RECITMP
      Do While Not .EOF
           SUINGRE# = XVALO(!DEBE)
           RETIRO# = XVALO(!HABER)
           FEMI% = CVINT(!FECHEMISI)
           REFE$ = SAFETEXT$(!REFERENCIA)
           CUO_ORIGINALX# = XVALO(!CUOTA_ORIGINAL)
           LGJO& = XVALO(!LEGAJO)
           DOCUM$ = SAFETEXT$(!CODICOM_LAR)
           SUINGREACUM# = SUINGREACUM# + SUINGRE#
           RETIROACUM# = RETIROACUM# + RETIRO#
           '
'           CNRT = CANTREGISTROS&("CAJABANC_COL", CONDI$, "NOMESS")
'           MSF.Rows = CNRT + 1
'           MSF.Row = 1
'           MSF.Col = 0
'
'           MSF.RowSel = MSF.Rows - 1
'           MSF.ColSel = MSF.Cols - 1
'
'           MSF.Clip = RECITMP.GetString(adClipString, -1, Chr(9), Chr(13), vbNullString)
'           MSF.Row = 1
          
           FILA& = FILA& + 1
           TXTCANTIFILAS = FILA&: TXTCANTIFILAS.Refresh
           MSF.Rows = FILA& + 1
           MSF.TextMatrix(FILA&, 1) = FECHATEX$(FEMI%)
           MSF.TextMatrix(FILA&, 2) = DOCUM$
           MSF.TextMatrix(FILA&, 3) = REFE$
           MSF.TextMatrix(FILA&, 4) = TRIM$(Str$(LGJO&))
           'MSF.TextMatrix(FILA&, 5) = VALOADMIN("ALUMNOS", "NOMBRE_APELLIDO", "LEGAJO = " & LGJO&, "NOMESS")
           MSF.TextMatrix(FILA&, 5) = SAFETEXT$(!NOMBRE_APELLIDO)
           MSF.TextMatrix(FILA&, 6) = TRIM$(FORMATNUM$(CUO_ORIGINALX#, "F12.2"))
           MSF.TextMatrix(FILA&, 7) = TRIM$(FORMATNUM$(SUINGRE#, "F12.2"))
           MSF.TextMatrix(FILA&, 8) = TRIM$(FORMATNUM$(RETIRO#, "F12.2"))
           .MoveNext
      Loop
    End With
    RECITMP.Close
    Set RECITMP = Nothing
    '
90  Call ENCABEZA_GRILLA

    Call SUMATOTAL(SUINGREACUM#, RETIROACUM#)
    MSF.Redraw = True

    'CARGO A PIE DE LA GRILLA EL TOTOL
    MSF.Rows = FILA& + 6
    MSF.TextMatrix(FILA& + 2, 6) = "Ingreso:": Call COLOREAR_GRILLA_SOLO_UNA_AFILA(CAJA.MSF, &HFCE4A7, FILA& + 2)
    MSF.TextMatrix(FILA& + 3, 6) = "Retiro:": Call COLOREAR_GRILLA_SOLO_UNA_AFILA(CAJA.MSF, &HFCE4A7, FILA& + 3)
    MSF.TextMatrix(FILA& + 4, 6) = "Resultado:": Call COLOREAR_GRILLA_SOLO_UNA_AFILA(CAJA.MSF, &HFCE4A7, FILA& + 4)
    '
    MSF.TextMatrix(FILA& + 2, 7) = Text15(0)
    MSF.TextMatrix(FILA& + 3, 8) = Text15(1)
    'SI ES POSITIVO VA DEL LADO DE INGRESO SI ES NEGATIVO VA DEL LADO RETIRO
    POS% = 8
    If XVALO(Text15(2)) > 0 Then POS% = 7
    MSF.TextMatrix(FILA& + 4, POS%) = Text15(2)
    '
    Screen.MousePointer = 1

End Sub
Sub SUMATOTAL(SUPAGOACUM#, RETIROACUM#)
    Text15(0) = FORMATNUM$(SUPAGOACUM#, "F14.2")
    Text15(1) = FORMATNUM$(RETIROACUM#, "F14.2")
    Text15(2) = FORMATNUM$(SUPAGOACUM# - RETIROACUM#, "F14.2")
    
End Sub
Sub ENCABEZA_GRILLA()
    MSF.Cols = 9
    NUCOLU% = 1: TEPRUEBA = "1": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 1: MSF.TextMatrix(0, NUCOLU% - 1) = ""
    NUCOLU% = 2: TEPRUEBA = "10/10/2012": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 1: MSF.TextMatrix(0, NUCOLU% - 1) = "Fecha Em."
    NUCOLU% = 3: TEPRUEBA = "12345678900000": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 1: MSF.TextMatrix(0, NUCOLU% - 1) = "Documento"
    NUCOLU% = 4: TEPRUEBA = "1234567890123456789012345678901234567890443": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 1: MSF.TextMatrix(0, NUCOLU% - 1) = "Referencia"
    NUCOLU% = 5: TEPRUEBA = "ABCDEf": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 6: MSF.TextMatrix(0, NUCOLU% - 1) = "Lgjo."
    NUCOLU% = 6: TEPRUEBA = "NOMBRE Y APELLIDO 0000000000": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 1: MSF.TextMatrix(0, NUCOLU% - 1) = "Nombre y Apellido"
    NUCOLU% = 7: TEPRUEBA = "12345678903": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 6: MSF.TextMatrix(0, NUCOLU% - 1) = "Imp.Original"
    NUCOLU% = 8: TEPRUEBA = "12345678901": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 6: MSF.TextMatrix(0, NUCOLU% - 1) = "Ingreso"
    NUCOLU% = 9: TEPRUEBA = "1234567890333": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 6: MSF.TextMatrix(0, NUCOLU% - 1) = "Retiro"
    '
    MSF.MergeCells = flexMergeRestrictColumns
    MSF.MergeCol(5) = True

End Sub

Sub LIMPIASUBTOTALES()

   For i% = 0 To 2
      Text15(i%) = ""
   Next i%
End Sub


Private Sub TEXT2_LostFocus(Index As Integer)
    Text2(Index) = FORMATNUM$(XVALO(Text2(Index)), "F14.2")
End Sub

