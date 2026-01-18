VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form INFO_GRILLAJM 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Valorización"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12375
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   12375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   9000
      Picture         =   "INFO_GRILLAJM.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Impresión de Grilla"
      Top             =   6080
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1600
   End
   Begin MSFlexGridLib.MSFlexGrid MSF2 
      Height          =   5715
      Left            =   2880
      TabIndex        =   6
      ToolTipText     =   "Escape para Ocultar"
      Top             =   0
      Visible         =   0   'False
      Width           =   6275
      _ExtentX        =   11060
      _ExtentY        =   10081
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   12632319
      BackColorBkg    =   16777215
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   105
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   4
      Top             =   735
      Visible         =   0   'False
      Width           =   1170
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5460
      Top             =   2415
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command8 
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
      Left            =   10740
      TabIndex        =   3
      Top             =   6080
      Width           =   1600
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   6045
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Doble Click Para Ver Detalle"
      Top             =   0
      Width           =   12345
      _ExtentX        =   21775
      _ExtentY        =   10663
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      AllowUserResizing=   1
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "INFO_GRILLAJM.frx":066A
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "INFO_GRILLAJM.frx":089E
      TabIndex        =   5
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   105
      TabIndex        =   2
      Top             =   105
      Visible         =   0   'False
      Width           =   3480
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   600
      TabIndex        =   1
      Top             =   7200
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "INFO_GRILLAJM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public NPROV%

Private Sub Command7_Click()

   'FECHAS DESDE Y HASTA TOMADAS DEL CAPTION
   '***************************************
   POS1% = 7 + InStr(INFO_GRILLAJM.Caption, "Desde:")
   FE1$ = Mid$(INFO_GRILLAJM.Caption, POS1%, 8)
   DESDEX% = FECHANUM(FE1$)
   '
   POS2% = 7 + InStr(INFO_GRILLAJM.Caption, "Hasta:")
   FE2$ = Mid$(INFO_GRILLAJM.Caption, POS2%, 8)
   '
   HASTAX% = FECHANUM(FE2$)
   On Error Resume Next
   Call ANACOSI09.PRINTINFOGRILLAJM(NPROV%, DESDEX%, HASTAX%)
   If Err Then
      On Error GoTo 0
      Call AMASTOK07.PRINTINFOGRILLAJM(NPROV%, DESDEX%, HASTAX%)
   End If
   On Error GoTo 0

   
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   
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
      Call GENERAXCEL(ARCHIEX$)
   End If
   
99 Command8.Enabled = True

End Sub
Sub GENERAXCEL(NombreDeArchivo$)
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
   For i& = 1 To MSF.Rows
       RENG1& = i&
       For COLU1& = 1 To MSF.Cols
          VALORX = MSF.TextMatrix(i& - 1, COLU1 - 1)
          If IsNumeric(VALORX) = True Then
             If InStr(VALORX, ".") > 0 Then
                VALO$ = VALORX
                VALO$ = REPLACAR(VALO$, ".", ",")
                VALORX = CDbl(VALO$)
             End If
          End If

          '
          XLSHEET.Cells(RENG1&, COLU1&) = VALORX
       Next COLU1&
   Next i&
   '
   On Error GoTo 30
   XLSHEET.SaveAs NombreDeArchivo$

   XLAPP.Visible = True
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



Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then msf2.Visible = False
End Sub

Private Sub Form_Load()
    Call APLICACIONSKINS(Me)

End Sub

Private Sub MSF_DblClick()
   'ASIGNO POSICION SELECCIONADA A VARIABLES
   i& = MSF.Row
   J& = MSF.COL
   '
   If i& < 1 Then Exit Sub
   '
   If J& = 0 Then
    
       'VENTANA VERTICAL
       '****************
       msf2.Visible = True
    
       msf2.Rows = MSF.Cols
       msf2.Cols = 2
       'VALIDO
       If i& = 0 Or i& > MSF.Rows Then Exit Sub
       If J& > MSF.Cols Then Exit Sub
       '
       NUCOLU% = 1: TEPRUEBA = "123456789012345678902345678901234": msf2.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: msf2.ColAlignment(NUCOLU% - 1) = 1: msf2.TextMatrix(0, NUCOLU% - 1) = ""
       NUCOLU% = 2: TEPRUEBA = "123456789012345678902345678901234": msf2.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: msf2.ColAlignment(NUCOLU% - 1) = 1: msf2.TextMatrix(0, NUCOLU% - 1) = "Cliente"
       '
       'CARGA EL ENCABEZADO A MODO VERTICAL
       For h& = 0 To MSF.Cols - 1
          DATO1$ = MSF.TextMatrix(0, h&)
          msf2.TextMatrix(h&, 0) = TRIM$(DATO1$)
       Next h&
       '
       'CARGA LOS DATOS DEL REGISTRO SELECCIONADO
       For h& = 0 To MSF.Cols - 1
          DATO1$ = MSF.TextMatrix(i&, h&)
          msf2.TextMatrix(h&, 1) = TRIM$(DATO1$)
       Next h&
       If msf2.Rows < 24 Then msf2.Rows = 24
       Call COLOREAR_GRILLASQL(msf2, &HEFF1F1)
   End If
   '
   'ABRE VENTANA DE DETALLE
   CI1% = CODINT%(MSF.TextMatrix(i&, 1))
   If CI1% > 0 Then
      Call ESTAXART(CI1%, NPROV%)
   End If
   
   
End Sub
Sub ESTAXART(CI1%, NCX1%)
   
   If CI1% < 1 Then Exit Sub
   '
   '
   'FECHAS DESDE Y HASTA TOMADAS DEL CAPTION
   '***************************************
   POS1% = 7 + InStr(INFO_GRILLAJM.Caption, "Desde:")
   FE1$ = Mid$(INFO_GRILLAJM.Caption, POS1%, 8)
   DESDEX% = FECHANUM(FE1$)
   '
   POS2% = 7 + InStr(INFO_GRILLAJM.Caption, "Hasta:")
   FE2$ = Mid$(INFO_GRILLAJM.Caption, POS2%, 8)
   '
   HASTAX% = FECHANUM(FE2$)
   '
   DESFE = FETEXCOR1$(DESDEX%)
   HASFE = FETEXCOR1$(HASTAX%, "HASTA")
   '
   
   'VALIDO SI HAY NOVEDADES CON MAS DE UNA MONEDA
   CONDI$ = " Fech_Mod  >=  '" & DESFE & " '"
   CONDI$ = CONDI$ + " and Fech_Mod  <=  '" & HASFE & " '"
   CONDI$ = CONDI$ + "  AND Lista_Repos = 'COSTO'"
   CONDI$ = CONDI$ + "  AND CODINT = " & CI1%
   CONDI$ = CONDI$ + "  AND NumProv  = " & NCX1%
'   CONDI2$ = "  AND MONEDA <= 1"
'   condi3$ = "  AND MONEDA > 1"
'   If CantRegistros("NOVEDADES", CONDI$ + CONDI2$, "NOMESS") > 0 Then PESO% = 1
'   If CantRegistros("NOVEDADES", CONDI$ + condi3$, "NOMESS") > 0 Then DOL% = 1
'   'SI HAY MAS DE UNA MONEDA PREGUNTO CUAL QUIERE VER
'   If PESO% + DOL% > 1 Then
'      OPX% = COMALTER%("Se Han Detectado Cambios de Precios en 2 Tipos de Monedas\\Ver Cambios en Pesos\Ver Cambios en Dolares")
'      If OPX% < 1 Or OPX% > 2 Then Exit Sub
'      TIPOMONE% = OPX%
'   ' SI HAY SOLO UNA MONEDA MUESTRA DIRECTAMENTE ESTA.
'   Else
'      TIPOMONE% = 1
'      If DOL% > 0 Then TIPOMONE% = 2
'   End If
   SQLX$ = "SELECT  * FROM NOVEDADES WITH(NOLOCK) "
   SQLX$ = SQLX$ + " WHERE   " & CONDI$
   If TIPOMONE% = 1 Then SQLX$ = SQLX$ + " AND MONEDA <= " & TIPOMONE%
   If TIPOMONE% = 2 Then SQLX$ = SQLX$ + " AND MONEDA >= " & TIPOMONE%
   SQLX$ = SQLX$ + " ORDER BY FECH_MOD"
   '
   INFO_GRILLA07.MSF.Redraw = False
   On Error Resume Next
   ANACOSI09.CARGA_ENCABEZADO_X_ART
   If Err Then
      On Error GoTo 0
      AMASTOK07.CARGA_ENCABEZADO_X_ART
   End If
   On Error GoTo 0
   
   REG& = 0
   Screen.MousePointer = 11
   Set RS = READSET(SQLX$)
   With RS
    Do While Not .EOF
      REG& = REG& + 1
      '
      INFO_GRILLA07.MSF.Rows = REG& + 1
      Call TRACE(24, REG&, FIN&)
      '
      INFO_GRILLA07.MSF.TextMatrix(REG&, 1) = FECHATEX$(CVINT(!Fech_Mod))
      P_ANT# = XVALO(!Prec_Ant)
      P_NEW# = XVALO(!Prec_New)
      EPAC$ = TRIM$(UCase$(EMPREAC$))
      'PARA JYM MUESTRA SIEMPRE EN LA MONEDA QUE FUE CARGADO (COMO SIEMPRE SE GUARDA EN $ EN EL CASO Q SEAN DOLARES LO PASA A DOLARES)
      If InStr(EPAC$, "J Y M") > 0 Or InStr(EPAC$, "CONTEPLAST") > 0 Then
        If (XVALO(!Moneda)) > 1 Then
           COTIZAX = 1
           If XVALO(!COTIZA_DOLAR) > 1 Then COTIZAX = XVALO(!COTIZA_DOLAR)
           P_NEW# = P_NEW# / COTIZAX
           P_ANT# = P_ANT# / COTIZAX
        End If
      End If
      INFO_GRILLA07.MSF.TextMatrix(REG&, 2) = FORMATNUM$(P_ANT#, "F12.2")
      INFO_GRILLA07.MSF.TextMatrix(REG&, 3) = FORMATNUM$(P_NEW#, "F12.2")
       PJE = ((P_NEW# / P_ANT#) - 1) * 100
      INFO_GRILLA07.MSF.TextMatrix(REG&, 4) = FORMATNUM$(PJE, "F8.2")
      INFO_GRILLA07.MSF.TextMatrix(REG&, 5) = FORMATNUM$(XVALO(!Moneda), "F2.0")
       
      INFO_GRILLA07.MSF.TextMatrix(REG&, 6) = TRIM$(ECOARCH$("FLEXUSR", MKI$(XVALO(!Nume_Usuar))))
      INFO_GRILLA07.MSF.TextMatrix(REG&, 7) = SAFETEXT$(!Observacion)
      IMPO_MONE_ORIGINAL = 0
      COTIZAX = 1
      If XVALO(!COTIZA_DOLAR) > 1 Then COTIZAX = XVALO(!COTIZA_DOLAR)
      If (XVALO(!Moneda)) > 1 Then IMPO_MONE_ORIGINAL = P_NEW# / COTIZAX
      If InStr(EPAC$, "J Y M") < 1 Then
        INFO_GRILLA07.MSF.TextMatrix(REG&, 8) = FORMATNUM$(IMPO_MONE_ORIGINAL, "F12.2")
        INFO_GRILLA07.MSF.TextMatrix(REG&, 9) = FORMATNUM$(COTIZAX, "F12.2")
      End If
      .MoveNext
    Loop
   End With
   
   RS.Close
   Set RS = Nothing
   Screen.MousePointer = 1
   '
   If REG& < 1 Then
     Call COMUNI("No se Detectaron Cambios en el Rango Seleccionado")
     Exit Sub
   End If
   
   If TIPOMONE% < 1 Then TIPOMONE% = 1
   DEMONE$ = DEMONECO$(TIPOMONE%)
   
   INFO_GRILLA07.MSF.Redraw = True
   INFO_GRILLA07.Caption = "Detalles de Cambios de Precios en " & DEMONE$ & "  - Art.: " & CODEXT$(CI1%) & " - Desde: " & FECHATEX$(DESDEX%) & " - Hasta: " & FECHATEX$(HASTAX%)
   Call COLOREAR_GRILLASQL(INFO_GRILLA07.MSF, &HEFF1F1)
   INFO_GRILLA07.Command7.Visible = True
   INFO_GRILLA07.CI1X = CI1%

   INFO_GRILLA07.Show 1
End Sub




