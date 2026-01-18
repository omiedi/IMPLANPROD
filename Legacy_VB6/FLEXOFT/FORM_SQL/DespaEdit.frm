VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form DespaEdit 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   $"DespaEdit.frx":0000
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   17535
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleWidth      =   17535
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TXTFECHA 
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
      Height          =   380
      Left            =   12250
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   240
      Width           =   1560
   End
   Begin VB.TextBox TXTNROPLANILLA 
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
      Height          =   380
      Left            =   13830
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   240
      Width           =   1995
   End
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
      Height          =   380
      Left            =   15840
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   240
      Width           =   1560
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Imprimir"
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
      Left            =   16000
      TabIndex        =   14
      Top             =   7440
      Width           =   1380
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
      Height          =   285
      Index           =   0
      Left            =   15840
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Text            =   "Cant. Bultos"
      Top             =   0
      Width           =   1560
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   6795
      Left            =   0
      TabIndex        =   5
      Top             =   600
      Width           =   17445
      _ExtentX        =   30771
      _ExtentY        =   11986
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
   Begin VB.TextBox Text5 
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
      Height          =   285
      Index           =   2
      Left            =   13830
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   "N° Planilla Despacho"
      Top             =   0
      Width           =   1995
   End
   Begin VB.TextBox Text5 
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
      Height          =   285
      Index           =   3
      Left            =   12250
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "Fecha"
      Top             =   0
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
      Left            =   13560
      TabIndex        =   7
      Top             =   7440
      Visible         =   0   'False
      Width           =   1380
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Height          =   30
      Left            =   2040
      TabIndex        =   4
      Top             =   1800
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   53
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   30
      Left            =   6720
      TabIndex        =   3
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
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "DespaEdit.frx":0099
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   120
      OleObjectBlob   =   "DespaEdit.frx":02CD
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1485
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
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
      WordWrap        =   -1  'True
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
      TabIndex        =   2
      Top             =   900
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Menu mnuArchivo1 
      Caption         =   "Archivo"
      Visible         =   0   'False
      Begin VB.Menu G1 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Selección de Planilla "
      Begin VB.Menu mnuMenu1 
         Caption         =   "MENU1"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMenu2 
         Caption         =   "MENU2"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSelecPlanDesp 
         Caption         =   "Seleccionar Planillas x Día y Despacho"
      End
      Begin VB.Menu s1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTodasXDia 
         Caption         =   "Seleccionar Todas las Planillas de un Día "
      End
   End
End
Attribute VB_Name = "DespaEdit"
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
    Campos$ = "Fecha/D8;Pedido/F8;Cta./F7;R.Social/A48;Desp.Nro./F10;N°Bulto/F7"
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    FRMZELECHO.Caption = TituloLista$
    
    Dim obj As New RECORXET
    TABLA$ = "DETABULTOS DB INNER JOIN DEUDOR12 D WITH(NOLOCK) ON DB.NROCLIE = D.Nume_Cli"
    Campos1$ = " CONVERT(varchar,feremovi,103)AS FECHA, NROPED,NroClie,D.RASO_CLI,NRODESPA ,count(numerobulto) AS total_bultos"
    CONDI$ = " STATUS = 1  GROUP BY nroped,nrodespa,CONVERT(varchar,feremovi,103),DB.NROCLIE,D.RASO_CLI ORDER BY DB.NROPED DESC"
    Set rs1 = obj.RS_GENERAL(TABLA$, Campos1$, CONDI$)

    Call Carga_MSF_Recordset(rs1, Campos$, FRMZELECHO.MSF2, 1)
    On Error Resume Next
    Set obj = Nothing
    rs1.Close
    Set rs1 = Nothing
    On Error GoTo 0
        
    FRMZELECHO.Width = 12000
50  FRMZELECHO.Show 1
    If XVALO(RESP_ZELE_VECT(5)) > 0 Then
       NroDespacho& = XVALO(RESP_ZELE_VECT(5))
    End If
    SeleccionarDespachoAHP& = NroDespacho&
    

End Function
Function SeleccionarDespachoConRemitoAHP&(retNroPed&, Optional DiaSelec%)
10  NroDespacho& = 0
    NROPEDret& = 0
    TituloLista$ = "Despachos Realizados (Doble Clic - Para Seleccionar)"
    If DiaSelec% > 33 Then TituloLista$ = "Despachos Realizados día: " & FECHATEX$(DiaSelec%) & " (Doble Clic - Para Seleccionar)"
    Campos$ = "Fecha/D8;Pedido/F7;Cta./F7;R.Social/A32;Desp.Nro./F7;N°Bulto/F7;Remitos/A54;N°de Guia/A12"
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    FRMZELECHO.Caption = TituloLista$
    
    Dim obj As New RECORXET
    TABLA$ = "DETABULTOS DB INNER JOIN DEUDOR12 D WITH(NOLOCK) ON DB.NROCLIE = D.Nume_Cli"
    Campos1$ = " CONVERT(varchar,feremovi,103)AS FECHA, NROPED,NroClie,D.RASO_CLI,NRODESPA ,max(numerobulto) AS total_bultos,nroguia"
    CONDI$ = " STATUS = 1  "
    If DiaSelec% > 33 Then 'SI VIENE CON FECHA FILTRA
       CONDI$ = CONDI$ + " AND feremovi >= '" & FETEXCOR1$(DiaSelec%) & "' AND  feremovi <= '" & FETEXCOR1$(DiaSelec%, "HASTA") & "'"
    End If
    CONDI$ = CONDI$ + " GROUP BY nroped,nrodespa,CONVERT(varchar,feremovi,103),DB.NROCLIE,D.RASO_CLI,nroguia ORDER BY DB.NROPED DESC"
    SQLX$ = " SELECT " & Campos1$
    SQLX$ = SQLX$ + " FROM " & TABLA$
    SQLX$ = SQLX$ + " WHERE " & CONDI$
    
    Set Rst = READSET(SQLX$)
    With Rst
        Do While Not .EOF
            FRMZELECHO.MSF2.Rows = FRMZELECHO.MSF2.Rows + 1
            FRMZELECHO.MSF2.TextMatrix(FRMZELECHO.MSF2.Rows - 1, 1) = SAFETEXT$(!FECHA)
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
        
    FRMZELECHO.Width = 16000
50  FRMZELECHO.Show 1
    If XVALO(RESP_ZELE_VECT(5)) > 0 Then
       NroDespacho& = XVALO(RESP_ZELE_VECT(5))
       NROPEDret& = XVALO(RESP_ZELE_VECT(2))
    End If
    SeleccionarDespachoConRemitoAHP& = NroDespacho&
    retNroPed& = NROPEDret&

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

 
   MSF.Rows = 1
99 Command1.Enabled = True

End Sub


Private Sub Command2_Click()
    Command2.Enabled = False
    RUTAEXCEL$ = Control("", "RULISDES")
    If EXISTE(RUTAEXCEL$) < 1 Then
       Call COMUNI("Falta Configurar Ruta de Plantilla Excel de Formulario de Despachos")
       GoTo 99
    End If
    
    IMPRESORA$ = TRIM$(Control$(IDENTER$, "PULISDES"))
    If IMPRESORA$ = "" Then PDES$ = Control$("*", "PULISDES")
    '
    Screen.MousePointer = 11
    'imprime los datos fijos
    Unload VENTABNEW
    With VENTABNEW.MSF
       .Rows = 5: .Cols = 2
       If XVALO(Me.TXTNROPLANILLA) < 1 Then
          .TextMatrix(1, 1) = "Despachos Varios"
       Else
          .TextMatrix(1, 1) = "N°: " & TRIM$(Str(XVALO(Me.TXTNROPLANILLA)))
       End If
         .TextMatrix(2, 1) = "Fecha: " & Me.TXTFECHA
         .TextMatrix(3, 1) = "Bultos: " & TRIM$(Str(XVALO(Me.TXTBULTOS)))
    End With
    
    If TRIM$(Control("BLOQEXCE", "BL0LISDE")) <> "SI" Then 'BLOQUEO LISTA DE DESPACHOS PARA NO ABRIR 2 AL MISMO TIEMPO
      Call GRACONTROL("BLOQEXCE", "BL0LISDE", "SI")
      'OJO SI SE LLAMA A PANTALLA A VENTABNEW CUANDO SE ACEPTA BORRA LAS FILAS QUE ESTAN VACIAS
      'VENTABNEW.Show 1
      Call Param_Print_Excel(RUTAEXCEL$, "HOJA1", VENTABNEW.MSF, IMPRESORA$, "1", 1, 20, "", 1, 1)                          'imprime en excel la 1,2 Y 3 columna de mi msf en la columna 1 a partir de la fila 11 en el excel (este graba E imprime)
      Unload VENTABNEW
'     A = FilaMaximaConContenidoDentroDeUnRangodeColumnas&("C:\Users\Omar\Desktop\Libro2.xlsx", "Hoja1", "C:J", 4)
     Call BorrarContenidoExcelXColumnas(RUTAEXCEL$, "Hoja1", "A5:L", 5)
     Call ExportarAExcel_MSF_IA(MSF, RUTAEXCEL$, "Hoja1", IMPRESORA$, 4, 0, "A1:L", 4)
     Call GRACONTROL("BLOQEXCE", "BL0LISDE", "NO")
    Else
      Call COMUNI("Archivo Bloqueado, Aguarde y Vuelva a Intentar")
      GoTo 99
    End If
99  Screen.MousePointer = 1
    Command2.Enabled = True
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
        Campos$ = Campos$ + ";Provincia/A0"          '5
        Campos$ = Campos$ + ";Provincia/A16"          '6
        Campos$ = Campos$ + ";R.Social Transporte/A20"  '7
        Campos$ = Campos$ + ";Domicilio/A20"          '8
        Campos$ = Campos$ + ";Localidad/A16"          '9
        Campos$ = Campos$ + ";Provincia/A0"          '10
        Campos$ = Campos$ + ";Provincia/A16"          '11
        Campos$ = Campos$ + ";Horario/A10"          '12
        Campos$ = Campos$ + ";Bulto/A0"               '13
        Campos$ = Campos$ + ";ID/A.0"               '14 de tabla detabultos
        Campos$ = Campos$ + ";Remito/A18"               '15
        Campos$ = Campos$ + ";N° de Guía/A14"               '16
        Campos$ = Campos$ + ";idDatoEntrega/A0"               '17 de tabla datoentrega
        
        
        CAMPOS_MSF$ = Campos$
End Function

Private Sub mnuAgregarGuia_Click()
10  NroDespacho& = SeleccionarDespachoConRemitoAHP&(retNroPed&)
    If NroDespacho& < 1 Then Exit Sub
    If retNroPed& < 1 Then Exit Sub
    NroGuia$ = LCase$(TRIM$(InputBox$(" Ingrese N° de Guía", "N° de Guía", NroGuia$)))
    If TRIM$(NroGuia$) = "" Then
       OPX% = COMALTER%("Atención No Se Ha Ingresado Ningún Dato !!! Desea Continuar \Si Confirma, el Campo Quedará en Blanco\Si Cancela, se Mantendrá el Valor Anterior\\No,Cancelar\Si, Continuar con la Modificación")
       If OPX% < 2 Or OPX% > 2 Then
       End If
    Else
       SQLX$ = "UPDATE DETABULTOS SET NroGuia ='" & NroGuia$ & "'"
       SQLX$ = SQLX$ + " WHERE NRODESPA = " & NroDespacho&
       SQLX$ = SQLX$ + " AND NROPED = " & retNroPed&
       FLEXCONN.Execute (SQLX$)
       Call COMUNI("N° de Guía Ingresado")
    End If

End Sub

Sub CargarDatosenGrillaSegunDespacho(NroDespacho&)
        MSF.Rows = 1
        CONDI$ = "D.NRODESPA= " & NroDespacho&
        CONDI$ = CONDI$ + " and Status =1"
        '
        SQLX$ = "SELECT D.NROPED,D.NumeroBulto,D.NROPED,DE.NROCLIE,DE.DOMI_SUC,DE.LOCA_SUC, DE.ProvEntre_ClI"
        SQLX$ = SQLX$ + " ,DE.DENO_TRA,DE.Deno_Tra,DOMI_TRA,DE.PROV_TRA,DE.HORATRANSPORTE,DE.Loca_Tra,DetaBultos_ID"
        SQLX$ = SQLX$ + " ,NroGuia,DatoEntrega_ID"
        SQLX$ = SQLX$ + " FROM DETABULTOS D WITH(NOLOCK)  "
        SQLX$ = SQLX$ + " INNER JOIN DATOENTREGA  DE ON D.NROPED = DE.NROPED  "
        SQLX$ = SQLX$ + " Where " & CONDI$
        SQLX$ = SQLX$ + " ORDER BY D.NROPED "
        Set rs = READSET(SQLX$)
        With rs
            Do While Not rs.EOF
                NROPED& = XVALO(!NROPED)
                RASOCLIENTE$ = rasocli$(XVALO(!NROCLIE))
                DOMICILIO_SUC$ = SAFETEXT(!DOMI_SUC)
'                NumeroBultoInterno% = XVALO(!NumeroBulto)
                LOCA_SUC$ = SAFETEXT(!LOCA_SUC)
                'PROVI_SUC$ = TRIM$(ECOARCH$("PROVINC", SAFETEXT(!ProvEntre_ClI)))
                PROVI_SUC$ = SAFETEXT(!ProvEntre_ClI)
                DesPROVI_SUC$ = TRIM$(ECOARCH$("PROVINC", SAFETEXT(!ProvEntre_ClI)))
                RSOCIAL_TRA$ = SAFETEXT(!DENO_TRA)
                DOMI_TRA$ = SAFETEXT(!DOMI_TRA)
                LOCA_TRA$ = SAFETEXT(!LOCA_TRA)
                'PROV_TRA$ = TRIM$(ECOARCH$("PROVINC", SAFETEXT(!PROV_TRA)))
                DesPROV_TRA$ = TRIM$(ECOARCH$("PROVINC", SAFETEXT(!PROV_TRA)))
                PROV_TRA$ = SAFETEXT(!PROV_TRA)
                HORA_TRA$ = SAFETEXT(!HoraTransporte)
                ID_DETABULTO& = XVALO(!DetaBultos_ID)
                
                MSF.Rows = MSF.Rows + 1
                MSF.Row = MSF.Rows - 1
                MSF.TextMatrix(MSF.Row, 1) = SAFETEXT$(NROPED&)
                MSF.TextMatrix(MSF.Row, 2) = RASOCLIENTE$
                MSF.TextMatrix(MSF.Row, 3) = DOMICILIO_SUC$
                MSF.TextMatrix(MSF.Row, 4) = LOCA_SUC$
                MSF.TextMatrix(MSF.Row, 5) = PROVI_SUC$
                MSF.TextMatrix(MSF.Row, 6) = DesPROVI_SUC$
                MSF.TextMatrix(MSF.Row, 7) = RSOCIAL_TRA$
                MSF.TextMatrix(MSF.Row, 8) = DOMI_TRA$
                MSF.TextMatrix(MSF.Row, 9) = LOCA_TRA$
                MSF.TextMatrix(MSF.Row, 10) = PROV_TRA$
                MSF.TextMatrix(MSF.Row, 11) = DesPROV_TRA$
                
                MSF.TextMatrix(MSF.Row, 12) = HORA_TRA$
                MSF.TextMatrix(MSF.Row, 13) = XVALO(!NumeroBulto)
                MSF.TextMatrix(MSF.Row, 14) = ID_DETABULTO&
                    CONDI$ = " NUPEDCLI = " & XVALO(!NROPED&) & " GROUP BY NUPEDCLI,DOPRILAR"
                MSF.TextMatrix(MSF.Row, 15) = StringAggFlex$("MOVISTO", CONDI$, "DOPRILAR", ",")
                MSF.TextMatrix(MSF.Row, 16) = SAFETEXT$(!NroGuia)
                MSF.TextMatrix(MSF.Row, 17) = XVALO(!DatoEntrega_ID)
                .MoveNext
            Loop
        End With
        rs.Close
        Set rs = Nothing
        '

        'RECORRO TODA LAS COLUMNAS DE LA FILA QUE SE AGREGO Y ADAPTO LA ALTURA AL TEXTO
        For REG& = 1 To MSF.Rows - 1
            For COL% = 1 To MSF.Cols - 1
                
                Call FRATEXTO(MSF.TextMatrix(REG&, COL%), (MSF.ColWidth(COL%) / 100))
                TXTX$ = ""
                For JJ& = 1 To CARETEX%
                  TXTX$ = TXTX$ + RETEX$(JJ&) + vbCrLf
                Next JJ&
                Label22 = TXTX$
                If UBound(RETEX$) > 1 Then
                   Call AdaptarAltoDeFilaSegunTexto(Label22, MSF, REG&, COL%)
                End If
            Next COL%
        Next REG&
        '
        'PINTO SI HAY REPETIDAS
'        Call PintarFilasRepetidas(MSF, 12)
        
        '
End Sub

Sub CargarDatosenGrillaSegunDia(DiaSelec%)
        MSF.Rows = 1
        CONDI$ = "Status =1"
        CONDI$ = CONDI$ + " and Feremovi >= '" & FETEXCOR1$(DiaSelec%) & "' AND Feremovi <='" & FETEXCOR1$(DiaSelec%, "Hasta") & "'"
        '
        SQLX$ = "SELECT D.NroDespa, D.NROPED,D.NumeroBulto,D.NROPED,DE.NROCLIE,DE.DOMI_SUC,DE.LOCA_SUC, DE.ProvEntre_ClI"
        SQLX$ = SQLX$ + " ,DE.DENO_TRA,DE.Deno_Tra,DOMI_TRA,DE.PROV_TRA,DE.HORATRANSPORTE,DE.Loca_Tra,DetaBultos_ID"
        SQLX$ = SQLX$ + " ,NroGuia,DatoEntrega_ID"
        SQLX$ = SQLX$ + " FROM DETABULTOS D WITH(NOLOCK)  "
        SQLX$ = SQLX$ + " INNER JOIN DATOENTREGA  DE ON D.NROPED = DE.NROPED  "
        SQLX$ = SQLX$ + " Where " & CONDI$
        SQLX$ = SQLX$ + " ORDER BY D.NRODESPA,D.NROPED,D.NumeroBulto "
        Set rs = READSET(SQLX$)
        With rs
            Do While Not rs.EOF
                NROPED& = XVALO(!NROPED)
                RASOCLIENTE$ = rasocli$(XVALO(!NROCLIE))
                DOMICILIO_SUC$ = SAFETEXT(!DOMI_SUC)
'                NumeroBultoInterno% = XVALO(!NumeroBulto)
                LOCA_SUC$ = SAFETEXT(!LOCA_SUC)
                'PROVI_SUC$ = TRIM$(ECOARCH$("PROVINC", SAFETEXT(!ProvEntre_ClI)))
                PROVI_SUC$ = SAFETEXT(!ProvEntre_ClI)
                DesPROVI_SUC$ = TRIM$(ECOARCH$("PROVINC", SAFETEXT(!ProvEntre_ClI)))
                RSOCIAL_TRA$ = SAFETEXT(!DENO_TRA)
                DOMI_TRA$ = SAFETEXT(!DOMI_TRA)
                LOCA_TRA$ = SAFETEXT(!LOCA_TRA)
                'PROV_TRA$ = TRIM$(ECOARCH$("PROVINC", SAFETEXT(!PROV_TRA)))
                DesPROV_TRA$ = TRIM$(ECOARCH$("PROVINC", SAFETEXT(!PROV_TRA)))
                PROV_TRA$ = SAFETEXT(!PROV_TRA)
                HORA_TRA$ = SAFETEXT(!HoraTransporte)
                ID_DETABULTO& = XVALO(!DetaBultos_ID)
                
                MSF.Rows = MSF.Rows + 1
                MSF.Row = MSF.Rows - 1
                MSF.TextMatrix(MSF.Row, 1) = SAFETEXT$(NROPED&) & " Desp. " & SAFETEXT$(!NroDespa)
                MSF.TextMatrix(MSF.Row, 2) = RASOCLIENTE$
                MSF.TextMatrix(MSF.Row, 3) = DOMICILIO_SUC$
                MSF.TextMatrix(MSF.Row, 4) = LOCA_SUC$
                MSF.TextMatrix(MSF.Row, 5) = PROVI_SUC$
                MSF.TextMatrix(MSF.Row, 6) = DesPROVI_SUC$
                MSF.TextMatrix(MSF.Row, 7) = RSOCIAL_TRA$
                MSF.TextMatrix(MSF.Row, 8) = DOMI_TRA$
                MSF.TextMatrix(MSF.Row, 9) = LOCA_TRA$
                MSF.TextMatrix(MSF.Row, 10) = PROV_TRA$
                MSF.TextMatrix(MSF.Row, 11) = DesPROV_TRA$
                
                MSF.TextMatrix(MSF.Row, 12) = HORA_TRA$
                MSF.TextMatrix(MSF.Row, 13) = XVALO(!NumeroBulto)
                MSF.TextMatrix(MSF.Row, 14) = ID_DETABULTO&
                    CONDI$ = " NUPEDCLI = " & XVALO(!NROPED&) & " GROUP BY NUPEDCLI,DOPRILAR"
                MSF.TextMatrix(MSF.Row, 15) = StringAggFlex$("MOVISTO", CONDI$, "DOPRILAR", ",")
                MSF.TextMatrix(MSF.Row, 16) = SAFETEXT$(!NroGuia)
                MSF.TextMatrix(MSF.Row, 17) = XVALO(!DatoEntrega_ID)
                .MoveNext
            Loop
        End With
        rs.Close
        Set rs = Nothing
        '

        'RECORRO TODA LAS COLUMNAS DE LA FILA QUE SE AGREGO Y ADAPTO LA ALTURA AL TEXTO
        For REG& = 1 To MSF.Rows - 1
            For COL% = 1 To MSF.Cols - 1
                
                Call FRATEXTO(MSF.TextMatrix(REG&, COL%), (MSF.ColWidth(COL%) / 100))
                TXTX$ = ""
                For JJ& = 1 To CARETEX%
                  TXTX$ = TXTX$ + RETEX$(JJ&) + vbCrLf
                Next JJ&
                Label22 = TXTX$
                If UBound(RETEX$) > 1 Then
                   Call AdaptarAltoDeFilaSegunTexto(Label22, MSF, REG&, COL%)
                End If
            Next COL%
        Next REG&
        '
        'PINTO SI HAY REPETIDAS
'        Call PintarFilasRepetidas(MSF, 12)
        
        '
End Sub

Private Sub mnuEliminarFila_Click()
   i& = 0
   REG& = MSF.MouseRow
   Call cmdEliminaItem(MSF, REG&)
   Call numerargrilla(MSF, 11)
   
   Exit Sub
   

   
End Sub

Private Sub mnuSelecPlanDesp_Click()
10 Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      feSelec% = FECHANUM(VALACT1$("SELFECHA"))
   Else
      Exit Sub
   End If

20 NroDespacho& = SeleccionarDespachoConRemitoAHP&(retNroPed&, feSelec%)
   If NroDespacho& < 1 Then GoTo 10
   TXTFECHA = FECHATEX$(feSelec%)
   TXTNROPLANILLA = NroDespacho&
   
  
End Sub

Private Sub mnuTodasXDia_Click()
   Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      feSelec% = FECHANUM(VALACT1$("SELFECHA"))
   Else
      Exit Sub
   End If
   Call CargarDatosenGrillaSegunDia(feSelec%)
   TXTFECHA = FECHATEX$(feSelec%)
   TXTNROPLANILLA = ""
   TXTBULTOS = MSF.Rows - 1
End Sub


Private Sub MSF_DblClick()
    Unload VENTABNEW
    Campos$ = CAMPOS_MSF$
    
    VENTANA.Inicializar = 0
    VENTANA.Campox = Campos$
    VENTANA.AgregaRegistro = 0
    VENTANA.UtilizaFormula = 1
    VENTANA.HabilitaBorrar = 1
    VENTANA.HabilitaInsertar = 0
    VENTANA.NoMostrarBotonTilde = 0
    VENTANA.TITULO = "Editar Despacho Activo"
    
    VENTANA.CampEditables = "3;4;5;7;8;9;10;16"
    VENTANA.ColEcoCampo(5) = "CODFLEX;NOMBRE;PROVINCIA"
    VENTANA.ColEcoCampo(10) = "CODFLEX;NOMBRE;PROVINCIA"
    VENTANA.Inicializar = 1
    '
    Call CopiarMsfAOtroMsf(Me.MSF, VENTABNEW.MSF, 1)
    '****************************************************************
    Call T_Espera(1 / 10) ' PARA QUE EL DOBLE CLICK NO APLIQUE SOBRE LA VENTANA QUE SE CARGA
30  VENTABNEW.Show 1
    If VENTABNEW.APROBADO% > 0 Then
       Call CopiarMsfAOtroMsf(VENTABNEW.MSF, Me.MSF, 1)
       For i& = 1 To MSF.Rows - 1
          NROPED& = XVALO(MSF.TextMatrix(i&, 1))
          Domicilio$ = Left$(MSF.TextMatrix(i&, 3), 128)
          localidad$ = Left$(MSF.TextMatrix(i&, 4), 128)
          Provincia$ = MSF.TextMatrix(i&, 5)
             DescProvCli$ = TRIM$(ECOARCH$("PROVINC", TRIM$(Provincia$)))
          RSocialTransporte$ = Left$(MSF.TextMatrix(i&, 7), 128)
          DomicilioTransp$ = Left$(MSF.TextMatrix(i&, 8), 128)
          localidadTransp$ = Left$(MSF.TextMatrix(i&, 9), 128)
          ProvinciaTransp$ = MSF.TextMatrix(i&, 10)
             DescProvTra$ = TRIM$(ECOARCH$("PROVINC", TRIM$(ProvinciaTransp$)))
          HorarioTransp$ = Left$(MSF.TextMatrix(i&, 12), 64)
          ID& = XVALO(MSF.TextMatrix(i&, 14))
          guia$ = MSF.TextMatrix(i&, 16)
          IdDatoEntrega& = XVALO(MSF.TextMatrix(i&, 17)) 'id de tabla datoentrega
          
          SQLX$ = "UPDATE DatoEntrega "
          SQLX$ = SQLX$ + "SET Domi_Suc = '" & Domicilio$ & "', "
          SQLX$ = SQLX$ + " Loca_Suc = '" & localidad$ & "', "
          SQLX$ = SQLX$ + " ProvEntre_Cli = '" & Provincia$ & "', "
          SQLX$ = SQLX$ + " Deno_Tra = '" & RSocialTransporte$ & "', "
          SQLX$ = SQLX$ + " Domi_Tra = '" & DomicilioTransp$ & "', "
          SQLX$ = SQLX$ + " Loca_Tra = '" & localidadTransp$ & "', "
          SQLX$ = SQLX$ + " Prov_Tra = '" & ProvinciaTransp$ & "', "
          SQLX$ = SQLX$ + " HoraTransporte = '" & HorarioTransp$ & "' "
          SQLX$ = SQLX$ + " WHERE NROPED = " & NROPED&
          SQLX$ = SQLX$ + " and DatoEntrega_ID = " & IdDatoEntrega&
          FLEXCONN.Execute (SQLX$)
 'TENER EN CUENTA QUE SI SE CARGAN DATOS A UN MISMO PEDIDO SE VAN A QUEDAR CON EL ULTIMO DATO DE LA ULTIMA FILA X PEDIDO
          
          'graba el nro de guia
          SQLX$ = "UPDATE DETABULTOS SET NroGuia = '" & guia$ & "'"
          SQLX$ = SQLX$ + " WHERE DETABULTOS_ID = " & ID&
          FLEXCONN.Execute (SQLX$)
'EN LA GUIA SI SE RESPETA PARA CADA FILA UN NUMERO DIFERENTE.
'        Campos$ = "Pedido/F8"           '1
'        Campos$ = Campos$ + ";R.Social Cliente/A12"          '2
'        Campos$ = Campos$ + ";Domicilio/A20"          '3
'        Campos$ = Campos$ + ";Localidad/A16"          '4
'        Campos$ = Campos$ + ";Provincia/A16"          '5
'        Campos$ = Campos$ + ";R.Social Transporte/A20"  '6
'        Campos$ = Campos$ + ";Domicilio/A20"          '7
'        Campos$ = Campos$ + ";Localidad/A16"          '8
'        Campos$ = Campos$ + ";Provincia/A16"          '9
'        Campos$ = Campos$ + ";Horario/A10"          '10
'        Campos$ = Campos$ + ";Bulto/A0"               '11
'        Campos$ = Campos$ + ";ID/A.0"               '12
'        Campos$ = Campos$ + "; Remito/A18"               '13
'        Campos$ = Campos$ + "; N° de Guía/A14"               '14

       Next i&
    End If
    
        'RECORRO TODA LAS COLUMNAS DE LA FILA QUE SE AGREGO Y ADAPTO LA ALTURA AL TEXTO
    For REG& = 1 To MSF.Rows - 1
    For COL% = 1 To MSF.Cols - 1
        Call FRATEXTO(MSF.TextMatrix(REG&, COL%), (MSF.ColWidth(COL%) / 100))
        TXTX$ = ""
        For JJ& = 1 To CARETEX%
          TXTX$ = TXTX$ + RETEX$(JJ&) + vbCrLf
        Next JJ&
        Label22 = TXTX$
        If UBound(RETEX$) > 1 Then
           Call AdaptarAltoDeFilaSegunTexto(Label22, MSF, REG&, COL%)
        End If
    Next COL%
   Next REG&
        '
' 'PINTO SI HAY REPETIDAS (CONSIDERO QUE AQUI NO HARIA FALTA MOSTRAR SI HAY REPETIDAS.)
' Call PintarFilasRepetidas(MSF, 12)

    Set VENTANA = Nothing
    Unload VENTABNEW
    NroDespa& = XVALO(TXTNROPLANILLA)
    TXTNROPLANILLA = 0
    TXTNROPLANILLA = NroDespa&
   'PRIMERO ORDENO POR LA COLUMNA DEL PROVEEDOR


End Sub




Private Sub TXTNROPLANILLA_Change()
    NroDespacho& = XVALO(TXTNROPLANILLA)
    If NroDespacho& < 1 Then
       Exit Sub
    End If
    Call CargarDatosenGrillaSegunDespacho(NroDespacho&)
    TXTBULTOS = MSF.Rows - 1
'    If XVALO(TXTNROPLANILLA) > 0 Then
'
End Sub


