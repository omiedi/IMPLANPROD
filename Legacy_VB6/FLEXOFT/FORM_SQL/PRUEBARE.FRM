VERSION 5.00
Object = "{D9AF33E0-7C55-11D5-9151-0000E856BC17}#1.0#0"; "fiscal010724.ocx"
Begin VB.Form PRUEBARE 
   BackColor       =   &H00E0E0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control de Impresión"
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00404040&
      Height          =   1320
      Left            =   3960
      ScaleHeight     =   1260
      ScaleWidth      =   870
      TabIndex        =   2
      Top             =   0
      Width           =   930
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "PRUEBARE.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Salir"
         Top             =   360
         Width           =   650
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0C0&
      Caption         =   "Prueba de Impresion de Remito"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3735
      Begin VB.CommandButton Command14 
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
         Height          =   630
         Left            =   840
         Picture         =   "PRUEBARE.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Impresion de Prueba de Remito"
         Top             =   360
         Width           =   2220
      End
   End
   Begin FiscalPrinterLibCtl.HASAR HASAR1 
      Left            =   0
      OleObjectBlob   =   "PRUEBARE.frx":0454
      Top             =   1200
   End
End
Attribute VB_Name = "PRUEBARE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Call FINAL("")
End Sub

Private Sub Command14_Click()
    HASAR1.Puerto = 1
    '
    HASAR1.Modelo = MODELO_P320             ' //Este OCX no es 100% para 715F
    '
    On Error Resume Next
    '
    HASAR1.Comenzar
    '
    If Err Then
      On Error GoTo 0
      TTXX$ = "Imposible Conectar con la Impresora Fiscal.\Revise el Encendido y las Conexiones\o Re-Configure para Impresora Común.\  \En Caso de Continuar, los Comandos se Presentarán\en Pantalla, pero NO se Enviarán a la Impresora."
      If COMALTER%(TTXX$ + "\\Configurar\Continuar") <> 2 Then
         Exit Sub
      End If
    Else
      Call MENSERR(24, "Conexion con Impresora Establecida")
    End If
    '
    HASAR1.PrecioBase = False
    '
    HASAR1.TratarDeCancelarTodo
    '
    On Error GoTo impresora_apag
           
Procesar:
    '
    On Error Resume Next
    'NCX% = XVALO(FORFACPRINT.Text1) ' **** VOLVER A ADAPTAR
    NCX% = 3061
    CUITT$ = SOLONUM$(CUITCLI$(NCX%))
    On Error GoTo 0
    TIDOCUX = TIPO_CUIT
    If PIVACLI%(NCX%) < 1 Then TIDOCUX = TIPO_DNI
    '
    PVCL% = PIVACLI%(NCX%)
    RESPONIVA = RESPONSABLE_INSCRIPTO
    If PVCL% < 1 Then RESPONIVA = CONSUMIDOR_FINAL
    If PVCL% = 2 Then RESPONIVA = RESPONSABLE_NO_INSCRIPTO
    If PVCL% = 3 Then RESPONIVA = NO_RESPONSABLE
    If PVCL% = 4 Then RESPONIVA = RESPONSABLE_EXENTO
    If PVCL% = 5 Then RESPONIVA = NO_CATEGORIZADO
    If PVCL% = 6 Then RESPONIVA = MONOTRIBUTO
    '
    ' ENCABEZADO
    On Error Resume Next
    HASAR1.Encabezado(1) = "VIAMONTE 2255"
    HASAR1.Encabezado(2) = "C1056ABI - CAPITAL FEDERAL"
      TELE$ = "Tel: 5411-4953-6114/6111"
    HASAR1.Encabezado(3&) = TELE$
      'CTE$ = TRIM$(FORFACPRINT.Text1) & " - " & TRIM$(FORFACPRINT.Label26)
      CTE$ = TRIM$(RASOCLI$(NCX%))
    HASAR1.Encabezado(4&) = "Cliente: " & TRIM$(FORMATNUM$(NCX%, "F5.0"))
    HASAR1.Encabezado(5&) = ""
    On Error GoTo 0
    '
    ' *** DATOS DEL CLIENTE ****
    HASAR1.DatosCliente "Cliente: " + TRIM$(RASOCLI$(NCX%)), CUITT$, TIDOCUX, RESPONIVA, "Domicilio: " + TRIM$(DOMICLI$(NCX%)) + "," + "CP " + TRIM$(LOCACLI$(NCX%))
    'Call MUESTRACOMANDO("DatosCliente", "Cliente: " + TRIM$(FORFACPRINT.Text2) & "(" & NCX% & ")", CUITT$, TIDOCUX, RESPONIVA)
    'If MUECOMHAS% < 2 Then HASAR1.DatosCliente "Cliente: " + TRIM$(FORFACPRINT.Text2), CUITT$, TIDOCUX, RESPONIVA, "Domicilio: " + TRIM$(FORFACPRINT.Text3) + "," + "CP " + TRIM$(FORFACPRINT.Text4)
    'IGUALMENTE SOLO IMPRIME LA LEYENDA CONSUMIDOR FINAL
    '
    ' *** ABRE COMPROBANTE NO HOMOLOGADO ***
    HASAR1.AbrirComprobanteNoFiscalHomologado REMITO
    
    'Call MUESTRACOMANDO("AbrirComprobanteFiscal", FACTURA_B)
    'If MUECOMHAS% < 2 Then HASAR1.AbrirComprobanteFiscal FACTURA_B
    '
    'Call VERIFICANUMEROREMITO()
    '
    HASAR1.ImprimirItemEnRemito "1000005579" & "  " & "TERMOCEL INTELLIGENT", 1
    HASAR1.ImprimirItemEnRemito "S/N: 557912120005", 0
    HASAR1.ImprimirItemEnRemito "1000005653" & "  " & "ALTA FRECUENCIA 6 ELECTRODOS", 2
    HASAR1.ImprimirItemEnRemito "S/N: 565301120017", 0
    
'    FIN& = ULTREG&("!CARDEFAC")
'    For U& = 1 To FIN&
'        XXX$ = REGLEIDO$("!CARDEFAC", U&)
'        CI% = CVI(Mid$(XXX$, 5, 2))
'        DESCR$ = TRIM$(Mid$(XXX$, 7, 64))
'        CANTIDA1# = ROUND(CVS(Left$(XXX$, 4)), 1)
'        PREUNI# = CVS(Mid$(XXX$, 429, 4))
'        ALICUIVA# = ROUND(CVS(Mid$(XXX$, 441, 4)), 2)
'        PRECONIVA# = ROUND(PREUNI# * (1 + ALICUIVA# / 100))
'        ACUMUIMPBRU = ACUMUIMPBRU + ROUND(PRECONIVA# * CANTIDA1#)
'        '
'        Call MUESTRACOMANDO("ImprimirItem", DESCR$, CANTIDA1#, PRECONIVA#, ALICUIVA#, 0)
'        If MUECOMHAS% < 2 Then HASAR1.ImprimirItem DESCR$, CANTIDA1#, PRECONIVA#, ALICUIVA#, 0
'        '
'    Next U&
    '
    HASAR1.CerrarComprobanteNoFiscalHomologado
    '
    HASAR1.ReimprimirComprobante
    HASAR1.ReimprimirComprobante
    
'    Call MUESTRACOMANDO("CerrarComprobanteFiscal")
'    If MUECOMHAS% < 2 Then HASAR1.CerrarComprobanteFiscal
    '
    Exit Sub

impresora_apag:

    If MsgBox("Error Impresora:" & Err.Description, vbRetryCancel, "Errores") = vbRetry Then
        Resume Procesar
    End If
   
    
'    On Error GoTo 0
'    '
'    Dim PRIREMI$(65) ' LECTURA DEL ARCHIVO TXT
'    NUAR% = FILEOPEN%(LUDAT$ & "\REMITO.TXT", 1, 128)
'    '
'    I& = 0
'    Do While Not EOF(NUAR%)
'      I& = I& + 1
'      Line Input #NUAR%, PRIREMI$(I&)
'    Loop
'    '
'    Call REPLA(PRIREMI$(4), "001-1500", 80, 12)
'    Call REPLA(PRIREMI$(7), "20/12/08 - 12:00", 83, 16)
'    Call REPLA(PRIREMI$(14), "1125", 10, 6)
'    Call REPLA(PRIREMI$(16), "Carlos Piedrabuena", 1, 32)
'    Call REPLA(PRIREMI$(16), "20-14011954-8", 79, 13)
'    Call REPLA(PRIREMI$(17), "Lisandro de la Torre 3343 Caseros Pcia.Buenos Aires", 1, 64)
'    Call REPLA(PRIREMI$(18), "Responsable Inscripto", 1, 32)
'    '
'    VALORRRR# = 2
'    Call REPLA(PRIREMI$(27), FORMATNUM(VALORRRR#, "F8.1"), 1, 8)
'
'    Call REPLA(PRIREMI$(27), "100005626", 13, 15)
'    Call REPLA(PRIREMI$(27), " * TENS ELECTROESTIMULADOR 2 CANALES ", 30, 48)
'    Call REPLA(PRIREMI$(28), "56209008866", 13, 12)
'    '
'    Call REPLA(PRIREMI$(29), FORMATNUM(VALORRRR#, "F8.1"), 1, 8)
'    Call REPLA(PRIREMI$(29), "100005627", 13, 15)
'    Call REPLA(PRIREMI$(29), " * TENS ELECTROESTIMULADOR 3 CANALES ", 30, 48)
'    Call REPLA(PRIREMI$(30), "56209008867", 13, 12)
'    '
'    Call REPLA(PRIREMI$(31), FORMATNUM(VALORRRR#, "F8.1"), 1, 8)
'    Call REPLA(PRIREMI$(31), "100005628", 13, 15)
'    Call REPLA(PRIREMI$(31), " * TENS ELECTROESTIMULADOR 4 CANALES ", 30, 48)
'    Call REPLA(PRIREMI$(32), "56209008868", 13, 12)
'
'    On Error GoTo controlerror
'    '
'10  HASAR1.AbrirComprobanteNoFiscal
'    'NX% = FILEOPEN%("X:\MEDITEA_BETA\FLEXOFT\VARIOS\REMITO11.TXT", 2, 128)
'
'    U& = 0
'    For U& = 1 To I&
'      HASAR1.ImprimirTextoNoFiscal (PRIREMI(I&))
'      'Print #NX%, PRIREMI$(U&)
'    Next U&
'    HASAR1.CerrarComprobanteNoFiscal
'    'Close #NX%
'    OPP% = COMALTER%("Repetir Impresion\\Si\No")
'    If OPP% = 1 Then
'        GoTo 10
'      ElseIf OPP% = 2 Then
'        GoTo 20
'    End If
'20  Call MENSERR(24, "Favor de Verificar la Impresion del Remito")
'    On Error GoTo 0
'    Exit Sub
'
'controlerror:
'    MsgBox Err.Number & "  -  " & Err.Description
'    Resume Next
End Sub


Private Sub Form_Load()
   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then Call FINAL("")

End Sub

Sub VERIFICANUMEROREMITO()
    '
    On Error Resume Next
    PRONUMERO$ = HASAR1.Respuesta(3)
    If Err Then
       On Error GoTo 0
       ' CALL MENSERR(24,"Imposible Determinar el Número\del Próximo Comprobante Fiscal.\   \Cierre la Aplicación y Re-Intente.")
       If COMALTER%("Imposible Determinar el Número\del Próximo Comprobante Fiscal.\\Abortar\Continuar") <> 2 Then
          On Error Resume Next
          HASAR1.TratarDeCancelarTodo
          On Error GoTo 0
          Call FINAL("")
       End If
    End If
    On Error GoTo 0
    '
    If XVALO(PRONUMERO$) <> XVALO(FORFACPRINT.Text10) Then
       'Call MENSERR(24, "El número del Próximo Comprobante Fiscal '" + PRONUMERO$ + "'\Difiere del Seteado en la Base de Datos FLEXOFT.\   \Debe Aparear los Números para Continuar.\   \Cierre la Aplicación y Re-Intente.")
       If COMALTER%("El Número del Próximo Comprobante Fiscal '" + PRONUMERO$ + "'\Difiere del Seteado en la Base de Datos FLEXOFT.\\Abortar\Continuar") <> 2 Then
          On Error Resume Next
          HASAR1.TratarDeCancelarTodo
          On Error GoTo 0
          Call FINAL("")
       End If
    End If
    '
End Sub
