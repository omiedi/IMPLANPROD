VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form HISTAPRO 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Historial de Aprobaciones"
   ClientHeight    =   7080
   ClientLeft      =   1050
   ClientTop       =   855
   ClientWidth     =   13860
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   13860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSFlexGridLib.MSFlexGrid MSF3 
      Height          =   3195
      Left            =   120
      TabIndex        =   7
      Top             =   3795
      Width           =   13635
      _ExtentX        =   24051
      _ExtentY        =   5636
      _Version        =   393216
      Rows            =   26
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
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
   Begin MSFlexGridLib.MSFlexGrid MSF0 
      Height          =   2355
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   13635
      _ExtentX        =   24051
      _ExtentY        =   4154
      _Version        =   393216
      Rows            =   26
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
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
      TabIndex        =   6
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
      Left            =   2370
      TabIndex        =   5
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
      TabIndex        =   4
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
      OleObjectBlob   =   "HISTAPRO.frx":0000
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   3240
      OleObjectBlob   =   "HISTAPRO.frx":0234
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin MSFlexGridLib.MSFlexGrid MSF4 
      Height          =   1635
      Left            =   120
      TabIndex        =   8
      Top             =   8235
      Width           =   13635
      _ExtentX        =   24051
      _ExtentY        =   2884
      _Version        =   393216
      Rows            =   26
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
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
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Consumos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   120
      TabIndex        =   12
      Top             =   7920
      Width           =   13635
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Detalle de Aprobaciones"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   120
      TabIndex        =   11
      Top             =   3480
      Width           =   13635
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cierres de Orden de Fabricación Activa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Top             =   645
      Width           =   13620
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Orden de Fabricación"
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
      TabIndex        =   9
      Top             =   45
      Width           =   1095
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
End
Attribute VB_Name = "HISTAPRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim TextoBuscarEntregados$
'Dim TextoBuscarPendientes$



'    SQLX$ = " SELECT M.FAMILIA, SUM(CANAPRO) AS CANP, SUM (CANPROY) AS CAPROY"
'    SQLX$ = SQLX$ + " FROM MASTER M WITH(NOLOCK) INNER JOIN ORDEFABR O WITH(NOLOCK)"
'    SQLX$ = SQLX$ + " ON M.CODINT = O.COD_INTE"
'    SQLX$ = SQLX$ + " AND  O.DESTINO ='" & TRIM$(Me.TXTDESTINO) & "'"
'    SQLX$ = SQLX$ + " AND (O.StatuOrf <= 3)"
'    SQLX$ = SQLX$ + " GROUP BY M.FAMILIA"
'    SQLX$ = SQLX$ + " ORDER BY M.FAMILIA"
'    Set RS = READSET(SQLX$)
'    With RS
'        Do While Not .EOF
'          TAFAM$ = SAFETEXT$(!FAMILIA)
'          FAMIX$ = ECOARCH$("TAFAMIL", TAFAM$)
'          CANTIAP = XVALO(!CANP)
'          CANPRO = XVALO(!CAPROY)
'          CANTI = XVALO(!CAPROY)
'          j& = j& + 1
'          MSF4.Rows = j& + 1
'          MSF4.TextMatrix(j&, 1) = FAMIX$
'          MSF4.TextMatrix(j&, 2) = FORMATNUM$(CANTI, "F12.1")
'          MSF4.TextMatrix(j&, 3) = FORMATNUM$(CANTI - CANTIAP, "F12.1")
'          MSF4.TextMatrix(j&, 4) = FORMATNUM$(CANTIAP, "F12.1")
'          MSF4.TextMatrix(j&, 5) = FORMATNUM$(CANTIAP * 100 / CANTI, "F12.1")
'          ACUMU2 = ACUMU2 + XVALO(MSF4.TextMatrix(j&, 2))
'          ACUMU3 = ACUMU3 + XVALO(MSF4.TextMatrix(j&, 3))
'          ACUMU4 = ACUMU4 + XVALO(MSF4.TextMatrix(j&, 4))
'          ACUMU5 = ACUMU5 + XVALO(MSF4.TextMatrix(j&, 5))
'          .MoveNext
'        Loop
'    End With
'    RS.Close
'    Set RS = Nothing
'    MSF4.Rows = MSF4.Rows + 2
'    MSF4.TextMatrix(MSF4.Rows - 1, 1) = "Totales:"
'    MSF4.TextMatrix(MSF4.Rows - 1, 2) = FORMATNUM$(ACUMU2, "F12.1")
'    MSF4.TextMatrix(MSF4.Rows - 1, 3) = FORMATNUM$(ACUMU3, "F12.1")
'    MSF4.TextMatrix(MSF4.Rows - 1, 4) = FORMATNUM$(ACUMU4, "F12.1")
'    MSF4.TextMatrix(MSF4.Rows - 1, 5) = FORMATNUM$(ACUMU5, "F12.1")
'    Call COLOREAR_GRILLA_SOLO_UNA_AFILA(Me.MSF4, &HE0E0E0, MSF4.Rows - 1)
'End Sub

Private Sub TXTBUSCAR_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      TXTBUSCAR.TEXT = ""
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
      For j& = 1 To GRID.Cols - 1
        If (InStr(UCase$(GRID.TextMatrix(i&, j&)), UCase$(TXTBUSCAR))) Then
          If (TRIM$(CMBFILTRO.TEXT)) = "" Or TRIM$(UCase$(CMBFILTRO)) = TRIM$(UCase$(GRID.TextMatrix(0, j&))) Then
            REG& = REG& + 1
            MSF_2.Rows = REG& + 1
            For H& = 1 To MSF_2.Cols - 1
              MSF_2.TextMatrix(REG&, H&) = GRID.TextMatrix(i&, H&)
            Next H&
          End If
          If (TRIM$(CMBFILTRO.TEXT)) = "" Then Exit For
        End If
      Next j&
    Next i&
    '
    MSF_2.Redraw = True
    
    '
End Sub
Private Sub CMBFILTRO_Change()
   TXTBUSCAR_Change
End Sub






Private Sub Form_Load()
      Campos$ = "CI/F0;Identif/A12;Código/A24;Descripción/A24;Q.Cerrada./F12.1;"
      Call CARGA_ENCABEZADO(Me.MSF0, Campos$, Me)
      
      Campos$ = "CI/F0;Código/A24;Descripción/A24;Q.Aprobada/F12.1;IDTEX/A16"
      Call CARGA_ENCABEZADO(Me.MSF3, Campos$, Me)
'     Text1 = FECHATEX$(HOY(HO$))
'     '
     Campos$ = "CI/F0;Código/A20;Descripción/A32;Consumo/F10.1;Idsituacion/A0;Deposito/F6"
     Call CARGA_ENCABEZADO(Me.MSF4, Campos$, Me)

End Sub

Private Sub TXTDATO1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      TXTDATO2.SetFocus
   End If
End Sub


Private Sub TXTDATO2_CHANGE()
   
End Sub
Private Sub Command6_Click()
   Campos$ = "CI/F0;Ord.Fabr./A12;Código/A24;Descripción/A24;Estado/A12;Q.Original/F12.1;Referencia/A0"
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO, , , 10000)
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO, , , 10000)
   '
   VectArticulos.cargarvector
   
   FRMZELECHO.Caption = "Ordenes de  Fabricación "
   FRMZELECHO.Width = 12500
   '
   FRMZELECHO.MSF1.Rows = 1
   Dim ESTAT(9)
   ESTAT(0) = "ABIERTA"
   ESTAT(3) = "CERRRADA"
   ESTAT(9) = "ANULADA"
   '
   Screen.MousePointer = 11
   j& = 0
   '*** CARGO LA LISTA DE SELECCION DESDE EL ARCHIVO ORFAPEN CON LA OF PENDIENTES DE APROBACION
   '¡¡ACA CARGAR SEGUN LA TABLA SQL DE PENDIENTES LOS SALDOS PARA GENERAR LA LISTA
16 SQLX$ = "SELECT O.COD_INTE,O.IDSUBOR,O.COD_EXTE, O.StatuOrf,O.CANPROY,o.referen from ORDEFABR O WITH(NOLOCK)INNER JOIN ORFAPEN F ON O.IDSUBOR = F.DOPRICOR"
   SQLX$ = SQLX$ + " WHERE NumeOrFa > 0 and f.codimovi = 'of' and f.idsituacion = 'GenCierre'"
   SQLX$ = SQLX$ + " ORDER BY NUMEORFA DESC"
   FRMZELECHO.MSF2.Redraw = False
   Set RS = READSET(SQLX$)
   With RS
   Do While Not .EOF
     Call TRACE(20, U&, FIN&)
     CIII% = XVALO(!COD_INTE)
     DOPRI$ = SAFETEXT$(!IdSubOr)
     CODI$ = SAFETEXT$(!Cod_Exte)
     DESXX$ = VectArticulos.descripcion(CIII%)
     ESTAT1% = XVALO(!StatuOrf)
     ESTATUX$ = ESTAT(ESTAT1%)
     CANTI1$ = TRIM$(FORMATNUM$(XVALO(!CANPROY), "f12.1"))
           
     
      j& = j& + 1
     FRMZELECHO.MSF2.Rows = j& + 1
     FRMZELECHO.MSF2.TextMatrix(j&, 1) = CIII%
     FRMZELECHO.MSF2.TextMatrix(j&, 2) = DOPRI$
     FRMZELECHO.MSF2.TextMatrix(j&, 3) = CODI$
     FRMZELECHO.MSF2.TextMatrix(j&, 4) = DESXX$
     FRMZELECHO.MSF2.TextMatrix(j&, 5) = ESTATUX$
     FRMZELECHO.MSF2.TextMatrix(j&, 6) = CANTI1
     FRMZELECHO.MSF2.TextMatrix(j&, 7) = SAFETEXT$(!Referen)
     
20   .MoveNext
   Loop
   End With
   RS.Close
   Set RS = Nothing
  '
   Screen.MousePointer = 1
   FRMZELECHO.MSF2.Redraw = True
   FRMZELECHO.Show 1
   NORFA$ = RESP_ZELE_VECT(2)
   If NORFA$ <> "" And NORFA$ <> TXTDESTINO Then
     TXTDESTINO = RESP_ZELE_VECT(2)
     TXTECODESTINO = RESP_ZELE_VECT(7)
     Call CARGAMSF0(TXTDESTINO)
   End If
   
End Sub


Private Sub MSF0_DblClick()
   REG& = MSF0.MouseRow
   If MSF0.Rows < 2 Then Exit Sub
   '
   Call CARGAMSF3(MSF0.TextMatrix(REG&, 2), TXTDESTINO)
   MSF4.Rows = 1
End Sub

Private Sub MSF0_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 114 Then
        If TextoBuscarEntregados$ <> "" Then
            FilaComienzo& = MSF0.Row
            If FilaComienzo& = MSF0.Rows - 1 Then FilaComienzo& = 1
20          For FilaGrid& = FilaComienzo& + 1 To MSF0.Rows - 1
                For ColumnaGrid& = 1 To MSF0.Cols - 1
                    MSF0.COL = ColumnaGrid&
                    If MSF0.Width > 0 Then
                        TextoCelda$ = SAFETEXT$(MSF0.TextMatrix(FilaGrid&, ColumnaGrid&))
                        If InStr(1, UCase(TextoCelda$), UCase(TextoBuscarEntregados$)) > 0 Then
                            Call SELECCIONARFILA(MSF0, FilaGrid&)
                            GoTo 99
                        End If
                    End If
                Next ColumnaGrid&
            Next FilaGrid&
            If PrimeraVuelta% = 0 Then
                PrimeraVuelta% = 1
                FilaComienzo& = 0
                GoTo 20
            End If
        End If
    End If
    
99  Exit Sub

End Sub
Sub CARGAMSF3(IDT$, NORFA$)
   MSF3.Rows = 1
16 SQLX$ = "SELECT IDTEXT,DOPRICOR, CANTSALID, CODIMOVI,COD_INTE,IDTEXT  FROM  ORFAPEN WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE IDTEXT = '" & IDT$ & "' AND IDSITUACION = 'APROCIERRE' AND DOPRICOR = '" & TRIM$(NORFA$) & "'"
   SQLX$ = SQLX$ + "  AND CODIMOVI = 'OF' ORDER BY FEREMOVI "
   Set RS = READSET(SQLX$)
   With RS
   Do While Not .EOF
     Call TRACE(20, U&, FIN&)
     CIII% = XVALO(!COD_INTE)
     DOPRI$ = SAFETEXT$(!DoPriCor)
     CODI$ = VectArticulos.codigo(CIII%)
     DESCRI$ = VectArticulos.descripcion(CIII%)
     CANTIX$ = FORMATNUM$(XVALO(!CANTSALID), "F12.1")
     ITX$ = SAFETEXT$(!IDTEXT)
      j& = j& + 1

     MSF3.Rows = j& + 1
     MSF3.TextMatrix(j&, 1) = CIII%
     MSF3.TextMatrix(j&, 2) = CODI$
     MSF3.TextMatrix(j&, 3) = DESCRI$
     MSF3.TextMatrix(j&, 4) = CANTIX$
     MSF3.TextMatrix(j&, 5) = ITX$
20   .MoveNext
   Loop
   End With
   RS.Close
   Set RS = Nothing
   
   If j& < 1 Then Call COMUNI("No Hay Aprobaciones Realizadas"): Exit Sub
   '
   Screen.MousePointer = 1
   '
End Sub

Sub CARGAMSF4(IDT$, NORFA$)
   MSF4.Rows = 1
16 SQLX$ = "SELECT IDTEXT,DOPRICOR, CANTINGRE, CODIMOVI,COD_INTE,IDTEXT,DEPOCONSU  FROM  ORFAPEN WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE IDTEXT = '" & IDT$ & "' AND IDSITUACION = 'APROCIERRE' AND DOPRICOR = '" & TRIM$(NORFA$) & "'"
   SQLX$ = SQLX$ + "  AND CODIMOVI = 'CF' ORDER BY FEREMOVI "
   Set RS = READSET(SQLX$)
   With RS
   Do While Not .EOF
     Call TRACE(20, U&, FIN&)
     CIII% = XVALO(!COD_INTE)
     DOPRI$ = SAFETEXT$(!DoPriCor)
     CODI$ = VectArticulos.codigo(CIII%)
     DESCRI$ = VectArticulos.descripcion(CIII%)
     CANTIX$ = FORMATNUM$(XVALO(!CANTINGRE), "F12.1") 'CANTINGRE POR QUE COMPENSA EL MOVIMIENTO GENERADO ORIGINALMENTE.
     ITX$ = SAFETEXT$(!IDTEXT)
     DEPOCONS% = XVALO(!DEPOCONSU)
      j& = j& + 1

     MSF4.Rows = j& + 1
     MSF4.TextMatrix(j&, 1) = CIII%
     MSF4.TextMatrix(j&, 2) = CODI$
     MSF4.TextMatrix(j&, 3) = DESCRI$
     MSF4.TextMatrix(j&, 4) = CANTIX$
     MSF4.TextMatrix(j&, 5) = ITX$
     MSF4.TextMatrix(j&, 6) = DEPOCONS%
     
20   .MoveNext
   Loop
   End With
   RS.Close
   Set RS = Nothing
   
   If j& < 1 Then Call COMUNI("No Hay O-Fabricación Para Autorizar"): Exit Sub
   '
   Screen.MousePointer = 1
   '
End Sub


Private Sub MSF3_DblClick()
'   REG& = MSF3.MouseRow
'   If MSF3.Rows < 2 Then Exit Sub
'   '
'
'   Call CARGAMSF4(MSF3.TextMatrix(REG&, 5), TXTDESTINO)

End Sub

Private Sub MSF3_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 114 Then
        If TextoBuscarPendientes$ <> "" Then
            FilaComienzo& = MSF3.Row
            If FilaComienzo& = MSF3.Rows - 1 Then FilaComienzo& = 1
20          For FilaGrid& = FilaComienzo& + 1 To MSF3.Rows - 1
                For ColumnaGrid& = 1 To MSF3.Cols - 1
                    MSF3.COL = ColumnaGrid&
                    If MSF3.Width > 0 Then
                        TextoCelda$ = SAFETEXT$(MSF3.TextMatrix(FilaGrid&, ColumnaGrid&))
                        If InStr(1, UCase(TextoCelda$), UCase(TextoBuscarPendientes$)) > 0 Then
                            Call SELECCIONARFILA(MSF3, FilaGrid&)
                            GoTo 99
                        End If
                    End If
                Next ColumnaGrid&
            Next FilaGrid&
            If PrimeraVuelta% = 0 Then
                PrimeraVuelta% = 1
                FilaComienzo& = 0
                GoTo 20
            End If
        End If
    End If
    
99  Exit Sub

End Sub




Private Sub TXTDESTINO_Change()
     MSF0.Rows = 1
     MSF3.Rows = 1
     MSF4.Rows = 1
'    Call CARGA_VENCIMIENTO_MSF1
'    Call CARGAMSF1
'    Call CARGAMSF2
'    Call CARGAMSF3
'    Call CARGA_VENCIMIENTO_MSF3
'    Call CARGATOTALES
End Sub
Sub CARGAMSF0(NORFA$)
   MSF0.Rows = 1
   '
   '
16 SQLX$ = "SELECT IDTEXT,DOPRICOR, CANTINGRE, CODIMOVI,COD_INTE  FROM  ORFAPEN WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE (CODIMOVI = 'OF') AND IDSITUACION = 'GenCierre' AND DOPRICOR = '" & TRIM$(NORFA$) & "'"
   SQLX$ = SQLX$ + " ORDER BY FEREMOVI "
   Set RS = READSET(SQLX$)
   With RS
   Do While Not .EOF
     Call TRACE(20, U&, FIN&)
     CIII% = XVALO(!COD_INTE)
     DOPRI$ = SAFETEXT$(!DoPriCor)
     CODI$ = VectArticulos.codigo(CIII%)
     DESCRI$ = VectArticulos.descripcion(CIII%)
     CANTIX$ = FORMATNUM$(XVALO(!CANTINGRE), "F12.1")
      j& = j& + 1

     MSF0.Rows = j& + 1
     MSF0.TextMatrix(j&, 1) = CIII%
     MSF0.TextMatrix(j&, 2) = SAFETEXT$(!IDTEXT)
     MSF0.TextMatrix(j&, 3) = CODI$
     MSF0.TextMatrix(j&, 4) = DESCRI$
     MSF0.TextMatrix(j&, 5) = CANTIX$
 
20   .MoveNext
   Loop
   End With
   RS.Close
   Set RS = Nothing
   
   '
   Screen.MousePointer = 1
   '

End Sub



