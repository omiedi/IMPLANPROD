VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form DespaCbrra 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pantalla de Carga Vía Lectora de Código QR "
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   17085
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7845
   ScaleWidth      =   17085
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TXTBULTOS 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   12240
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   7440
      Width           =   1560
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   0
      Left            =   10680
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   "Cant. Bultos"
      Top             =   7440
      Width           =   1560
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   15600
      TabIndex        =   10
      Top             =   7440
      Width           =   1500
   End
   Begin VB.TextBox TXTLECTURA 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1080
      TabIndex        =   0
      Text            =   " "
      Top             =   7440
      Width           =   3060
   End
   Begin VB.ListBox Lista 
      Height          =   2010
      Left            =   120
      TabIndex        =   6
      Top             =   5160
      Visible         =   0   'False
      Width           =   2775
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Height          =   30
      Left            =   2040
      TabIndex        =   5
      Top             =   1800
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   53
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   30
      Left            =   6720
      TabIndex        =   4
      Top             =   2640
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   53
      _Version        =   393216
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
      Left            =   1320
      ScaleHeight     =   165
      ScaleWidth      =   1110
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "DespaCbrra.frx":0000
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   120
      OleObjectBlob   =   "DespaCbrra.frx":0234
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   1485
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   7395
      Left            =   240
      TabIndex        =   8
      Top             =   0
      Width           =   16965
      _ExtentX        =   29924
      _ExtentY        =   13044
      _Version        =   393216
      Cols            =   5
      ForeColor       =   0
      WordWrap        =   -1  'True
      AllowBigSelection=   -1  'True
      Enabled         =   -1  'True
      FocusRect       =   2
      AllowUserResizing=   2
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
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   3600
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   ">>>>>>>>"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   120
      TabIndex        =   7
      Top             =   7515
      Width           =   975
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   0
      TabIndex        =   3
      Top             =   900
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Menu mnuArchivo1 
      Caption         =   "Archivo"
      Visible         =   0   'False
      Begin VB.Menu mnuEliminarFila 
         Caption         =   "Eliminar Fila Seleccionada"
         Visible         =   0   'False
      End
      Begin VB.Menu G1 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mnuMenu1 
         Caption         =   "MENU1"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMenu2 
         Caption         =   "MENU2"
         Visible         =   0   'False
      End
      Begin VB.Menu s1 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuUtilidades 
      Caption         =   "Utilidades"
      Begin VB.Menu mnuDespaReal 
         Caption         =   "Reingreso de Bultos Despachados"
      End
      Begin VB.Menu mnuEditarDespacho 
         Caption         =   "Editar Datos de Despacho"
      End
      Begin VB.Menu mnuAgregarGuia 
         Caption         =   "Agregar N° de Guía"
      End
   End
End
Attribute VB_Name = "DespaCbrra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FrmLeft, FrmTop, FrmWidth, FrmHeight As Long
Dim GridTop, GridLeft, GridWidth, GridHeight As Long
Dim FRAMETop, FRAMELeft, FRAMEWidth, FRAMEHeight As Long
Dim Com1Top, Com2Top, Com2Left As Long
Dim Com1LEFT As Long

Dim LecturaDatosOriginales As Byte
Public MENU1%
Public MENU2%
Public MENU3%
Public NOPINTAR%
Public SeleccionaFilaDobleClick%
Public SinEfectoClickEnGrilla%

Dim GLOBALFRM_strContenidoCelda$
Dim GLOBALFRM_FILA&

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Function SeleccionarDespachoAHP&()
10  NroDespacho& = 0
    TituloLista$ = "Despachos Realizados (Doble Clic - Para Seleccionar)"
    Campos$ = "Fecha/D8;Pedido/F8;Cta./F7;R.Social/A48;Desp.Nro./F10;N°Bulto/F6"
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    FRMZELECHO.Caption = TituloLista$
    
    Dim obj As New RECORXET
    TABLA$ = "DETABULTOS DB INNER JOIN DEUDOR12 D WITH(NOLOCK) ON DB.NROCLIE = D.Nume_Cli"
    Campos1$ = " feremovi, NROPED,NroClie,D.RASO_CLI,NRODESPA ,count(numerobulto) AS total_bultos"
    CONDI$ = " STATUS = 1  GROUP BY nroped,nrodespa,feremovi,DB.NROCLIE,D.RASO_CLI ORDER BY feremovi DESC,NRODESPA desc"
    Set RS1 = obj.RS_GENERAL(TABLA$, Campos1$, CONDI$)

    Call Carga_MSF_Recordset(RS1, Campos$, FRMZELECHO.MSF2, 1)
    On Error Resume Next
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    On Error GoTo 0
        
    FRMZELECHO.Width = 12000
50  FRMZELECHO.Show 1
    If XVALO(RESP_ZELE_VECT(5)) > 0 Then
       NroDespacho& = XVALO(RESP_ZELE_VECT(5))
    End If
    SeleccionarDespachoAHP& = NroDespacho&
    

End Function
Function SeleccionarDespachoConRemitoAHP&(retNroPed&, Optional retNroGuia$, Optional DiaSelec%)
'tener en cuenta que el nro de guia en este afecta a la combinacion del pedido y el nro de despacho
'en particular lo hace desde la pantalla de edicion de todo el despacho

10  NroDespacho& = 0
    NROPEDret& = 0
    NroGuia$ = ""
    TituloLista$ = "Despachos Realizados (Doble Clic - Para Seleccionar)"
    If DiaSelec% > 33 Then TituloLista$ = "Despachos Realizados día: " & FECHATEX$(DiaSelec%) & " (Doble Clic - Para Seleccionar)"
    
    Campos$ = "Fecha/D8;Pedido/F7;Cta./F7;R.Social/A32;Desp.Nro./F7;Bultos/F6;Remitos/A54;N°de Guia/A24"
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    FRMZELECHO.Caption = TituloLista$
    
    Dim obj As New RECORXET
    TABLA$ = "DETABULTOS DB INNER JOIN DEUDOR12 D WITH(NOLOCK) ON DB.NROCLIE = D.Nume_Cli"
    Campos1$ = " feremovi , NROPED,NroClie,D.RASO_CLI,NRODESPA ,max(numerobulto) AS total_bultos,nroguia"
    CONDI$ = " STATUS = 1  "
    If DiaSelec% > 33 Then 'SI VIENE CON FECHA FILTRA
       CONDI$ = CONDI$ + " AND feremovi >= '" & FETEXCOR1$(DiaSelec%) & "' AND  feremovi <= '" & FETEXCOR1$(DiaSelec%, "HASTA") & "'"
    End If
    CONDI$ = CONDI$ + " GROUP BY nroped,nrodespa,feremovi,DB.NROCLIE,D.RASO_CLI,nroguia ORDER BY feremovi desc,DB.NROPED DESC"
    SQLX$ = " SELECT " & Campos1$
    SQLX$ = SQLX$ + " FROM " & TABLA$
    SQLX$ = SQLX$ + " WHERE " & CONDI$
    Set Rst = READSET(SQLX$)
    With Rst
        Do While Not .EOF
            FRMZELECHO.MSF2.Rows = FRMZELECHO.MSF2.Rows + 1
            FRMZELECHO.MSF2.TextMatrix(FRMZELECHO.MSF2.Rows - 1, 1) = CVDAT(CVINT(!FeReMovi))
            FRMZELECHO.MSF2.TextMatrix(FRMZELECHO.MSF2.Rows - 1, 2) = XVALO(!NROPED)
            FRMZELECHO.MSF2.TextMatrix(FRMZELECHO.MSF2.Rows - 1, 3) = XVALO(!NROCLIE)
            FRMZELECHO.MSF2.TextMatrix(FRMZELECHO.MSF2.Rows - 1, 4) = SAFETEXT$(!raso_cli)
            FRMZELECHO.MSF2.TextMatrix(FRMZELECHO.MSF2.Rows - 1, 5) = XVALO(!NroDespa)
            FRMZELECHO.MSF2.TextMatrix(FRMZELECHO.MSF2.Rows - 1, 6) = XVALO(!total_bultos)
              CONDI$ = " NUPEDCLI = " & XVALO(!NROPED) & " GROUP BY NUPEDCLI,DOPRILAR"
            FRMZELECHO.MSF2.TextMatrix(FRMZELECHO.MSF2.Rows - 1, 7) = StringAggFlex$("MOVISTO", CONDI$, "DOPRILAR", ",")
            FRMZELECHO.MSF2.TextMatrix(FRMZELECHO.MSF2.Rows - 1, 8) = SAFETEXT$(!NroGuia)
            .MoveNext
        Loop
    End With
    On Error Resume Next
    Rst.Close
    Set Rst = Nothing
    On Error GoTo 0
        
    FRMZELECHO.Width = 17000
50  FRMZELECHO.Show 1
    If XVALO(RESP_ZELE_VECT(5)) > 0 Then
       NroDespacho& = XVALO(RESP_ZELE_VECT(5))
       NROPEDret& = XVALO(RESP_ZELE_VECT(2))
       NroGuia$ = RESP_ZELE_VECT(8)
    Else
      Exit Function
    End If
    SeleccionarDespachoConRemitoAHP& = NroDespacho&
    retNroPed& = NROPEDret&
    retNroGuia$ = NroGuia$

End Function

Private Sub Command1_Click()
   Command1.Enabled = False
   'NO VALIDO QUE NO HAYA FILAS REPETIDAS, POR QUE DE ULTIMA VA A GRABAR EN EL MISMO REGISTRTO POR QUE TIENEN EL MISMO ID
   'VEAMOS QUE QUIEREN HACER.
   cadenaID$ = ""
   For REG& = 1 To MSF.Rows - 1
      ID& = XVALO(MSF.TextMatrix(REG&, 12))
      If TRIM$(cadenaID$) <> "" Then
          cadenaID$ = cadenaID$ & "," & TRIM$(CStr(ID&))
      Else
          cadenaID$ = TRIM$(CStr(ID&))
      End If
   Next REG&
'
   nroDesp& = 1 + XVALO(VALOBADA("DETABULTOS", "TOP 1 NRODESPA", "NRODESPA > 0 ORDER BY NRODESPA DESC", "NOMESS"))
   CONDI$ = " DetaBultos_ID IN (" & cadenaID$ & ")"
   
   SQLX$ = " UPDATE DETABULTOS SET STATUS = 1 , NRODESPA = " & nroDesp&
   SQLX$ = SQLX$ + " , FeReMovi = '" & Now & "'"
   SQLX$ = SQLX$ + " WHERE " & CONDI$
   FLEXCONN.Execute (SQLX$)
   '
   Call COMUNI("PROCESO TERMINADO")
   '
   'Call cmdImprimir_MSF(MSF)
'   Call CARGA_TABLA_IMPRE(MSF, CAMPOS_MSF$, "DESPA_AHP", "Despachos AHP", "TABLA_IMPRE", RET_NAMECONS$, 1)
''   Call FINAL("?" & RET_NAMECONS$)
'   FILTX$ = ""
'   Call STDFORM("DSPA-AHP", FILTX$)
'
'   Call GenerarHTMLDesdeFlexGrid(MSF, "C:\FLEXOFT\Archivo.html", 0, Observa$)
''   VER TEMA DE LA IMPRESION OJO SI NO CONFIGURAR IMPRESORA PDF.
 
   MSF.Rows = 1
99 Command1.Enabled = True

End Sub
Private Sub Form_Load()
   
    Campos$ = CAMPOS_MSF$
    Call CREATABLA_VENTAS
    Call CARGA_ENCABEZADO(Me.MSF, Campos$, Me)


End Sub
Function CAMPOS_MSF$()
        Campos$ = "Pedido/F8"           '1
        Campos$ = Campos$ + ";R.Social Cliente/A12"          '2
        Campos$ = Campos$ + ";Domicilio/A20"          '3
        Campos$ = Campos$ + ";Localidad/A16"          '4
        Campos$ = Campos$ + ";Provincia/A16"          '5
        Campos$ = Campos$ + ";R.Social Transporte/A20"  '6
        Campos$ = Campos$ + ";Domicilio/A20"          '7
        Campos$ = Campos$ + ";Localidad/A16"          '8
        Campos$ = Campos$ + ";Provincia/A16"          '9
        Campos$ = Campos$ + ";Horario/A10"          '10
        Campos$ = Campos$ + ";Bulto/A0"               '11
        Campos$ = Campos$ + ";ID/A0"               '12
        
        CAMPOS_MSF$ = Campos$
End Function

Private Sub mnuAgregarGuia_Click()
   Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      DiaSelec% = FECHANUM(VALACT1$("SELFECHA"))
   Else
      Exit Sub
   End If

10   NroDespacho& = SeleccionarDespachoConRemitoAHP&(retNroPed&, retNroGuia$, DiaSelec%)

  'NroDespacho& = SeleccionarDespachoConRemitoAHP&(retNroPed&, retNroGuia$)
    If NroDespacho& < 1 Then Exit Sub
    If retNroPed& < 1 Then Exit Sub
    NroGuia$ = retNroGuia$
    NroGuia$ = LCase$(TRIM$(InputBox$(" Ingrese N° de Guía", "N° de Guía", NroGuia$)))
    If TRIM$(NroGuia$) = "" Then
       OPX% = COMALTER%("Atención No Se Ha Ingresado Ningún Dato !!! Desea Continuar \Si Confirma, el Campo Quedará en Blanco\Si Cancela, se Mantendrá el Valor Anterior\\No,Cancelar\Si, Continuar con la Modificación")
       If OPX% < 2 Or OPX% > 2 Then
          GoTo 10
       End If
    End If
    SQLX$ = "UPDATE DETABULTOS SET NroGuia ='" & Left$(NroGuia$, 48) & "'"
    SQLX$ = SQLX$ + " WHERE NRODESPA = " & NroDespacho&
    SQLX$ = SQLX$ + " AND NROPED = " & retNroPed&
    FLEXCONN.Execute (SQLX$)
    Call COMUNI("N° de Guía Ingresado")
    GoTo 10
    
    

End Sub

Private Sub mnuDespaReal_Click()
    
    
'10  NroDespacho& = 0
'    TituloLista$ = "Despachos Realizados (Doble Clic - Ver Detalle)"
'    Campos$ = "Fecha/D8;Pedido/F8;Cta./F7;R.Social/A48;Desp.Nro./F10;Bultos/F6"
'    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
'    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
'    FRMZELECHO.Caption = TituloLista$
'
'    Dim obj As New RECORXET
'    TABLA$ = "DETABULTOS DB INNER JOIN DEUDOR12 D WITH(NOLOCK) ON DB.NROCLIE = D.Nume_Cli"
'    Campos1$ = " CONVERT(varchar,feremovi,103)AS FECHA, NROPED,NroClie,D.RASO_CLI,NRODESPA ,count(numerobulto) AS total_bultos"
'    CONDI$ = " STATUS = 1  GROUP BY nroped,nrodespa,CONVERT(varchar,feremovi,103),DB.NROCLIE,D.RASO_CLI ORDER BY DB.NROPED DESC"
'    Set RS1 = obj.RS_GENERAL(TABLA$, Campos1$, CONDI$)
'
'    Call Carga_MSF_Recordset(RS1, Campos$, FRMZELECHO.MSF2, 1)
'    On Error Resume Next
'    Set obj = Nothing
'    RS1.Close
'    Set RS1 = Nothing
'    On Error GoTo 0
'
'    FRMZELECHO.Width = 12000
'50  FRMZELECHO.Show 1
'    If XVALO(RESP_ZELE_VECT(5)) > 0 Then
'       NroDespacho& = XVALO(RESP_ZELE_VECT(5))
'    End If
    
10  NroDespacho& = SeleccionarDespachoAHP&
    If NroDespacho& < 1 Then Exit Sub
    
    Call VerZelechoSegunNroDespacho(NroDespacho&)
    GoTo 10


End Sub
Sub VerZelechoSegunNroDespacho(NroDespacho&)
'   Reingreso de Bultos Despachados
    TituloLista$ = "Detalle de Despacho N°: " & NroDespacho&
    Campos$ = "N°Bulto/F8;Peso/F8.2;Pedido/F8;Cta./F7;R.Social/A48;Fecha/D8;ID/A0"
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    FRMZELECHO.Caption = TituloLista$
    
    Dim obj As New RECORXET
    TABLA$ = "DETABULTOS DB WITH(NOLOCK) INNER JOIN DEUDOR12 D ON D.Nume_Cli= DB.NROCLIE"
    Campos1$ = " numerobulto,pesokg, NROPED,NroClie,D.RASO_CLI,feremovi,DETABULTOS_ID"
    CONDI$ = " STATUS = 1 AND NRODESPA = " & NroDespacho& & "  ORDER BY DB.NROPED DESC"
    Set RS1 = obj.RS_GENERAL(TABLA$, Campos1$, CONDI$)

    Call Carga_MSF_Recordset(RS1, Campos$, FRMZELECHO.MSF2, 1)
    On Error Resume Next
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    On Error GoTo 0
        
    FRMZELECHO.Width = 12000
50  FRMZELECHO.Show 1
    If XVALO(RESP_ZELE_VECT(7)) > 0 Then
       ID& = XVALO(RESP_ZELE_VECT(7))
    End If
    
    If ID& < 1 Then Exit Sub
    SQLX$ = "UPDATE DETABULTOS SET STATUS = -1,NRODESPA = 0"
    SQLX$ = SQLX$ + " , FeReMovi = '" & Now & "'"
    SQLX$ = SQLX$ + " WHERE DETABULTOS_ID =" & ID&
    FLEXCONN.Execute (SQLX$)
    Call COMUNI("Se ha Reingresado el Bulto Seleccionado")
    
End Sub

Private Sub mnuEditarDespacho_Click()
   DespaEdit.Show 1
End Sub

Private Sub MSF_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
         REG& = MSF.MouseRow
         COL% = MSF.MouseCol
         Call MENU_CELDA(MSF, REG&, COL%)
      End If
End Sub

Sub MENU_CELDA(GRID As msFlexGrid, REG&, COL%)
        If GRID.Rows <= 1 Then Exit Sub
        
'        Call COLOREAR_GRILLA_SOLACELDA(GRID, vbYellow, REG&, Int(COL%))   ' LA PINTA DE AMARILLA
        Call INVICTRL
        GLOBALFRM_strContenidoCelda$ = TRIM$(GRID.TextMatrix(REG&, COL%))  'CELDA SELECCIONADA
        GLOBALFRM_FILA& = REG&

        PopupMenu mnuArchivo1
        Call INVICTRL
        'Call COLOREAR_GRILLA_SOLACELDA(GRID, vbWhite, REG&, COL%)  ' LA PINTA DE AMARILLA
End Sub

Sub INVICTRL()
  '
'  mnucopiar.Visible = Not (mnucopiar.Visible)
  mnuEliminarFila.Visible = Not (mnuEliminarFila.Visible)
  '
End Sub

Private Sub mnuEliminarFila_Click()
   i& = 0
   REG& = MSF.MouseRow
   Call cmdEliminaItem(MSF, REG&)
   Call numerargrilla(MSF, 11)
   
   Exit Sub
   
'   Do While i& < MSF.Rows - 1
'      i& = i& + 1
'      If TRIM$(MSF.TextMatrix(i&, 8)) = NROREMI$ Then
'       Call cmdEliminaItem(MSF, i&)
'       i& = i& - 1
'      End If
'   Loop
'   If MSF.Rows < 2 Then
'      Text3 = ""
'   Else
'      Call numerargrilla(MSF, 1)
'   End If
   
End Sub
Private Sub TXTLECTURA_KeyPress(KeyAscii As Integer)
     
     If KeyAscii <> 13 Then Exit Sub
     
     LecturaEtiqueta$ = TRIM$(TXTLECTURA)
     
     If LecturaEtiqueta$ = "" Then Exit Sub
     AscCar% = TieneCaracteresAsciiMayora127(LecturaEtiqueta$)
     If AscCar% > 0 Then
       LecturaEtiqueta$ = REPLACAR$(LecturaEtiqueta$, Chr$(AscCar%), "")
     End If
     LECTURA$ = LecturaEtiqueta$
     On Error Resume Next
     
     NROPED& = XVALO(Left$(LecturaEtiqueta$, 8))
     NroBulto% = XVALO(Mid$(LecturaEtiqueta$, 9, 4))
     On Error GoTo 0
     '
     Call CargarDatosenGrilla(NROPED&, NroBulto%)
     TXTBULTOS = MSF.Rows - 1
     TXTLECTURA = ""
     Call PINTATEXT(TXTLECTURA) 'lo selecciono por que le aparece una Ê luego que lee de este modo la proxima lectura lo sobreescribe
     
End Sub
Sub CargarDatosenGrilla(NROPED&, NroBulto%)
      
        RASOCLIENTE$ = ""
        DOMICILIO_SUC$ = ""
        LOCA_SUC$ = ""
        LOCA_SUC$ = ""
        RSOCIAL_TRA$ = ""
        DOMI_TRA$ = ""
        PROV_TRA$ = ""
        HORA_TRA$ = ""
        HAYDATOS% = 0
        CONDI$ = "D.NROPED= " & NROPED&
        CONDI$ = CONDI$ + " and D.NumeroBulto =" & NroBulto%
        
        'VALIDACION SI EL BULTO YA FUE ENTREGADO
        CONDI2$ = "NroPed = " & NROPED& & " And NumeroBulto = " & NroBulto% & " AND STATUS >0"
        yaEntregado& = CantRegistros("DETABULTOS", CONDI2$, "NOMESS")
        If yaEntregado& > 0 Then
           Call COMUNI("Atención !!! Imposible Agregar Bulto yá Entregado")
           Exit Sub
        End If
        '
        SQLX$ = "SELECT D.NROPED,D.NumeroBulto,D.NROPED,DE.NROCLIE,DE.DOMI_SUC,DE.LOCA_SUC, DE.ProvEntre_ClI"
        SQLX$ = SQLX$ + " ,DE.DENO_TRA,DE.Deno_Tra,DOMI_TRA,DE.PROV_TRA,DE.HORATRANSPORTE,DE.Loca_Tra,DetaBultos_ID"
        SQLX$ = SQLX$ + " FROM DETABULTOS D WITH(NOLOCK)  "
        SQLX$ = SQLX$ + " INNER JOIN DATOENTREGA  DE ON D.NROPED = DE.NROPED  "
        SQLX$ = SQLX$ + " Where " & CONDI$
        Set rs = READSET(SQLX$)
        With rs
            If Not rs.EOF Then
                HAYDATOS% = 1
                NROPED& = XVALO(!NROPED)
                RASOCLIENTE$ = rasocli$(XVALO(!NROCLIE))
                DOMICILIO_SUC$ = SAFETEXT(!DOMI_SUC)
'                NumeroBultoInterno% = XVALO(!NumeroBulto)
                LOCA_SUC$ = SAFETEXT(!LOCA_SUC)
                PROVI_SUC$ = TRIM$(ECOARCH$("PROVINC", SAFETEXT(!ProvEntre_ClI)))
                RSOCIAL_TRA$ = SAFETEXT(!DENO_TRA)
                DOMI_TRA$ = SAFETEXT(!DOMI_TRA)
                LOCA_TRA$ = SAFETEXT(!LOCA_TRA)
                PROV_TRA$ = TRIM$(ECOARCH$("PROVINC", SAFETEXT(!PROV_TRA)))
                HORA_TRA$ = SAFETEXT(!HoraTransporte)
                ID_DETABULTO& = XVALO(!DetaBultos_ID)
           End If
        End With
        rs.Close
        Set rs = Nothing
        '
        If HAYDATOS% < 1 Then Exit Sub
        
        
        MSF.Rows = MSF.Rows + 1
        MSF.Row = MSF.Rows - 1
        MSF.TextMatrix(MSF.Row, 1) = NROPED&
        MSF.TextMatrix(MSF.Row, 2) = RASOCLIENTE$
        MSF.TextMatrix(MSF.Row, 3) = DOMICILIO_SUC$
        MSF.TextMatrix(MSF.Row, 4) = LOCA_SUC$
        MSF.TextMatrix(MSF.Row, 5) = PROVI_SUC$
        MSF.TextMatrix(MSF.Row, 6) = RSOCIAL_TRA$
        MSF.TextMatrix(MSF.Row, 7) = DOMI_TRA$
        MSF.TextMatrix(MSF.Row, 8) = LOCA_TRA$
        MSF.TextMatrix(MSF.Row, 9) = PROV_TRA$
        MSF.TextMatrix(MSF.Row, 10) = HORA_TRA$
        MSF.TextMatrix(MSF.Row, 12) = ID_DETABULTO&

        'RECORRO TODA LAS COLUMNAS DE LA FILA QUE SE AGREGO Y ADAPTO LA ALTURA AL TEXTO
        For COL% = 1 To MSF.Cols - 1
            Call FRATEXTO(MSF.TextMatrix(MSF.Row, COL%), (MSF.ColWidth(COL%) / 100))
            TXTX$ = ""
            For JJ& = 1 To CARETEX%
              TXTX$ = TXTX$ + RETEX$(JJ&) + vbCrLf
            Next JJ&
            Label22 = TXTX$
            If UBound(RETEX$) > 1 Then
               Call AdaptarAltoDeFilaSegunTexto(Label22, MSF, MSF.Row, COL%)
            End If
        Next COL%
        '
        'PINTO SI HAY REPETIDAS
        Call PintarFilasRepetidas(MSF, 12)
        
        '
        Call numerargrilla(MSF, 11)
End Sub
