VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form TESTCAE09 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Solicitud de CAE Manual"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   6180
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   6180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrRecuperaComprobanteDesdeAfip 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   720
      Top             =   3000
   End
   Begin VB.Timer tmrRecuperaUltimoNroCbte 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   3000
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   27
      Top             =   420
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   4200
      Top             =   3000
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   4680
      Top             =   3000
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   840
      Top             =   1995
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   240
      Top             =   2160
   End
   Begin VB.TextBox Text3 
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
      Height          =   2580
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   3480
      Width           =   5010
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
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
      Height          =   300
      Left            =   1350
      TabIndex        =   7
      Top             =   2760
      Width           =   2235
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
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
      Height          =   300
      Left            =   1350
      TabIndex        =   6
      Top             =   2280
      Width           =   2480
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00624E28&
      ForeColor       =   &H80000008&
      Height          =   6255
      Left            =   5280
      ScaleHeight     =   6225
      ScaleWidth      =   1425
      TabIndex        =   0
      Top             =   0
      Width           =   1455
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   100
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   29
         Top             =   5760
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -120
            TabIndex        =   30
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sol. CAE con ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         TabIndex        =   23
         Top             =   2640
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command16 
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
         Height          =   690
         Left            =   120
         Picture         =   "TESTCAE09.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Re-Impresión de Comprobantes"
         Top             =   945
         Width           =   650
      End
      Begin VB.CommandButton Command22 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "TESTCAE09.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Proceso Automático de Solicitud de Cae Pendientes"
         Top             =   4680
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "TESTCAE09.frx":07B4
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   240
         Width           =   650
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   4200
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2D3C0&
      Height          =   1935
      Left            =   140
      TabIndex        =   9
      Top             =   120
      Width           =   5040
      Begin VB.CommandButton Command5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Este boton pide el CAE"
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
         Left            =   240
         TabIndex        =   24
         Top             =   1440
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Solicitud de CAE Manual"
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
         Left            =   1560
         TabIndex        =   19
         Top             =   1440
         Width           =   2415
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
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
         Left            =   1920
         TabIndex        =   18
         Top             =   960
         Width           =   2805
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "."
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
         Left            =   4560
         TabIndex        =   15
         Top             =   600
         Width           =   165
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Solicitar CAE"
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
         Left            =   2520
         TabIndex        =   14
         Top             =   1440
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
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
         Left            =   1920
         TabIndex        =   12
         Top             =   600
         Width           =   2730
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
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
         Left            =   1920
         TabIndex        =   11
         Top             =   240
         Width           =   570
      End
      Begin VB.Image Image1 
         Height          =   495
         Left            =   4200
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label MODOAUTO 
         Height          =   255
         Left            =   4320
         TabIndex        =   20
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label44 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Id Transacción"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -120
         TabIndex        =   17
         Top             =   690
         Width           =   1890
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Número Provisorio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         TabIndex        =   13
         Top             =   1040
         Width           =   1650
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Punto de Venta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   360
         TabIndex        =   10
         Top             =   320
         Width           =   1410
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "TESTCAE09.frx":08FE
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   60
      OleObjectBlob   =   "TESTCAE09.frx":0B32
      TabIndex        =   28
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Emisión"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   3720
      TabIndex        =   32
      Top             =   2880
      Width           =   540
   End
   Begin VB.Label lblFechaEmisionAfip 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   345
      Left            =   4320
      TabIndex        =   31
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   345
      Left            =   4320
      TabIndex        =   26
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Vto.:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   3920
      TabIndex        =   25
      Top             =   2400
      Width           =   420
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2280
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cae Devuelto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   -360
      TabIndex        =   4
      Top             =   2355
      Width           =   1560
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nº Factura Asignado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   -120
      TabIndex        =   3
      Top             =   2655
      Width           =   1320
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Respuesta Recibida"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   1770
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   5880
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu mnuTransacciones 
      Caption         =   "Transacciones"
      Begin VB.Menu mnuRecuperaComprobantes 
         Caption         =   "Recuperar Comprobantes desde Afip"
         Begin VB.Menu mnuFacturaElectronicaComun 
            Caption         =   "Factura Electrónica Común"
         End
         Begin VB.Menu mnuRecuperaFacturaCredito 
            Caption         =   "Factura Electronica Común de Crédito"
         End
         Begin VB.Menu mnuRecuperaCbteBienesDeCapital 
            Caption         =   "Factura de Bienes de Capital"
         End
         Begin VB.Menu mnuRecuperaFacturaExpo 
            Caption         =   "Factura de Exportación"
         End
      End
      Begin VB.Menu mnuSolicitudCaeManual 
         Caption         =   "Solicitud de CAE Manual"
         Visible         =   0   'False
         Begin VB.Menu mnuSolicitudFacturaComun 
            Caption         =   "Factura Electrónica Común"
         End
      End
   End
   Begin VB.Menu mnuConfiguracion 
      Caption         =   "Configuración"
      Visible         =   0   'False
      Begin VB.Menu mnuInicioFacturaElectronica 
         Caption         =   "Inicio Nueva Factura Electrónica"
      End
   End
End
Attribute VB_Name = "TESTCAE09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ESTADO%
Dim CONTAD&
Dim INDICE&
Dim MODO%
Dim CANTCOPIAS%
Public FILTX$
Dim ARCSQR$
Dim SOLOCAE%
Dim ORCOP$(9)
Dim FACTU$()
Public YAGRABOCAE%

Dim RecuperacionPuntoVenta%
Dim RecuperacionTipoComprobante%()
Dim RecuperacionUltimoNumeroPorTipo&()
Dim RecuperacionHuecos$()
Dim RecuperacionModoCredito As Boolean
Dim ModoRecuperacionDeComprobante As Boolean

Dim ModoSolicitudManual As Boolean

Sub RecuperaConsultaCae(IDENCAE$)

    Condi$ = " TEXTOCAE Is Not Null  AND IDCAENUM = " & XVALO(IDENCAE$)

    DATOX$ = VALOBADA("CAJABANC", "TEXTOCAE", Condi$, "NOMESS")
    If DATOX$ = "" Then
       Call COMUNI("No Se Han Grabado Correctamente Los Datos Para Solicitar el CAE a la A.F.I.P.\Capture esta Pantalla y Consulte a su Soporte de Sistemas")
       Exit Sub
    End If
    '
TXTRAZA$ = TXTRAZA$ & "PUNTO 58 = " & Now & vbCrLf
    'FECHA DE HOY
    HOII% = HOY(HOS$)
    FECH1$ = ARMAFECHCAE(HOII%) ' FORMATO AAAAMMDD

    FE% = CVINT(VALOBADA("CAJABANC", "FECHEMISI", Condi$, "NOMESS"))
    FECH2$ = ARMAFECHCAE(FE%) ' FORMATO AAAAMMDD
    If FECH1$ <> FECH2$ Then DATOX$ = REPLACAR(DATOX$, FECH2$, FECH1$)  'REEMPLAZO LA FECHA QUE TENIA POR LA DE HOY
    CARPETOK$ = "C:\FLEXOFT\FACELEC\"
    Archiv$ = CARPETOK$ + "CONSULTA_CAE_" + IDENCAE$ + ".TXT"
    NARCHI2% = FILEOPEN%(Archiv$, 2, 128) ' ABRE EL ARCHIVO DE ESCRITURA
    Print #NARCHI2%, DATOX$ ' ESCRIBE EN EL ARCHIVO EL NUEVO TEXTO
    Close #NARCHI2%

End Sub

Sub SolicitudManualFacturaElectronicaComun(PuntoVenta As Integer)

        RecuperacionPuntoVenta% = PuntoVenta
        ReDim RecuperacionTipoComprobante%(1)
        
        TipoCbte$ = Mid$(Text6, 1, 2) & Mid$(Text6, 18, 1)
        Select Case TipoCbte$
            Case "FCA": RecuperacionTipoComprobante%(0) = 1 'FCA
            Case "NDA": RecuperacionTipoComprobante%(0) = 2 'NDA
            Case "NCA": RecuperacionTipoComprobante%(0) = 3 'NCA
            Case "FCB": RecuperacionTipoComprobante%(0) = 6 'FCB
            Case "NDB": RecuperacionTipoComprobante%(0) = 7 'NDB
            Case "NCB": RecuperacionTipoComprobante%(0) = 8 'NCB
        End Select
        
        RecuperacionModoCredito = False
        ReDim RecuperacionUltimoNumeroPorTipo&(1)

        tmrRecuperaUltimoNroCbte.Enabled = True

End Sub

Sub SolicitudManualFacturaBienesDeCapital(PuntoVenta As Integer)

        RecuperacionPuntoVenta% = PuntoVenta
        ReDim RecuperacionTipoComprobante%(1)
        
        TipoCbte$ = Mid$(Text6, 1, 2) & Mid$(Text6, 18, 1)
        Select Case TipoCbte$
            Case "FCA": RecuperacionTipoComprobante%(0) = 1 'FCA
            Case "NDA": RecuperacionTipoComprobante%(0) = 2 'NDA
            Case "NCA": RecuperacionTipoComprobante%(0) = 3 'NCA
            Case "FCB": RecuperacionTipoComprobante%(0) = 6 'FCB
            Case "NDB": RecuperacionTipoComprobante%(0) = 7 'NDB
            Case "NCB": RecuperacionTipoComprobante%(0) = 8 'NCB
        End Select
        
        RecuperacionModoCredito = False
        ReDim RecuperacionUltimoNumeroPorTipo&(1)

        tmrRecuperaUltimoNroCbte.Enabled = True

End Sub


Sub SolicitudManualFacturaExportacion(PuntoVenta As Integer)

        RecuperacionPuntoVenta% = PuntoVenta
        ReDim RecuperacionTipoComprobante%(1)
        
        TipoCbte$ = Mid$(Text6, 1, 2) & Mid$(Text6, 18, 1)
        Select Case TipoCbte$
            Case "FCE": RecuperacionTipoComprobante%(0) = 19 'FCE
            Case "NDE": RecuperacionTipoComprobante%(0) = 20 'NDE
            Case "NCE": RecuperacionTipoComprobante%(0) = 21 'NCE
        End Select
        
        RecuperacionModoCredito = False
        ReDim RecuperacionUltimoNumeroPorTipo&(1)

        tmrRecuperaUltimoNroCbte.Enabled = True

End Sub

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Unload Me
End Sub

Sub Command4_Click()
    ' SOLOCAE% = 0
    Timer1.Enabled = False
        
    If ModoSolicitudManual = True Then Exit Sub
    
    Call COMUNI("Antes de Volver a Solicitar Autorización para el Comprobante Elegido,\el Sistema presentará una Lista de los Comprobantes\Autorizados en AFIP y NO Registrados en el Sistema.\ \En caso de Encontrar en dicha Lista el Comprobante sobre el cual Solicitó Autorización,\deberá Recuperar el mismo Haciendo un Doble Clic sobre el renglón correspondiente.")
    ModoSolicitudManual = True
    Command4.Enabled = False
    
    ' REVISA SI EN AFIP EXISTEN COMPROBANTES DEL MISMO PUNTO DE VENTA
    ' QUE NO ESTAN REGISTRADOS EN EL SISTEMA. LA IDEA ES QUE EL USARIO VERIFIQUE
    ' SI EN LISTA DE COMPROBANTES SE ENCUENTRA EL QUE ESTA QUERIENDO OBTENER CAE MANUALMENTE.
    PuntoVentaFacturaComun% = XVALO(CONTROL("FACTURA", "PUVENELE")) 'P VTA. NORMAL
    PuntoVentaFacturaCredito% = XVALO(CONTROL("FACTURA", "PUVECRED"))  'P VTA. CREDITO
    PuntoVentaBienesCapital% = XVALO(CONTROL("FACTURA", "PUVENCAP")) 'P.VTA. B.CAPITAL
    PuntoVentaExportacion% = XVALO(CONTROL("FACTURA", "PUVEFAEX")) 'P VTA. EXPORTACION
    
    Select Case XVALO(Text4)
        Case PuntoVentaFacturaComun%: Call SolicitudManualFacturaElectronicaComun(PuntoVentaFacturaComun%)
        Case PuntoVentaFacturaCredito%: Call SolicitudManualFacturaElectronicaComun(PuntoVentaFacturaCredito%)
        Case PuntoVentaBienesCapital%: Call SolicitudManualFacturaBienesDeCapital(PuntoVentaBienesCapital%)
        Case PuntoVentaExportacion%: Call SolicitudManualFacturaExportacion(PuntoVentaExportacion%)
    End Select
       
End Sub

Sub SOLICITACAE()
      EJECTUTA$ = "C:\FLEXOFT\FACELEC\"
TXTRAZA$ = TXTRAZA$ & "PUNTO 64 = " & Now & vbCrLf
      PVTA% = XVALO(Text4)
      IDENCAE$ = TRIM$(Text5)
      If PVTA% < 1 Or IDENCAE$ = "" Then
           If MODOAUTO = "NO" Then Call COMUNI("Imposible Solicitar CAE.\Datos Incompletos o Invalidos.")
           Exit Sub
         Else
           On Error Resume Next
           Kill EJECTUTA$ + "CAE_" + IDENCAE$ + ".TXT" ' ELIMINAMOS EL ARCHIVO
           On Error GoTo 0
           If EXISTE%(EJECTUTA$ + "CAE_" + IDENCAE$ + ".TXT") > 0 Then
              If MODOAUTO = "NO" Then Call MENSERR(24, "Imposible Solicitar CAE.\No se Pudo Eliminar Archivo '" + REPLACAR$(EJECTUTA$ + "CAE_" + IDENCAE$ + ".TXT'", "\", "/"))
              Exit Sub
           End If
           Call PIDECAE ' (PVTA%, IDENCAE$)
      End If
TXTRAZA$ = TXTRAZA$ & "PUNTO 65 = " & Now & vbCrLf
''      INDICE& = 1
''      ReDim FACTU$(INDICE&)
''      FACTU$(INDICE&) = IDENCAE$
      
''      ESTADO% = 0
''      CONTAD& = 0
''      Timer2.Enabled = True
      
End Sub

Sub PIDECAE() ' (PVTA%, ID$)
     '
     PVTA% = XVALO(Text4)
     IDENCAE$ = TRIM$(Text5)
     EJECTUTA$ = "C:\FLEXOFT\FACELEC\"
     '
     Call RecuperaConsultaCae(IDENCAE$)
     '
     COMPROB$ = Mid$(Text6, 1, 2)
     Variante$ = Mid$(Text6, 18, 1)
     CondicionSql$ = "codicom_lar='" & TRIM$(Text6) & "' and idcaenum=" & IDENCAE$ & " and modoopcimovim='cr'"
     EsModoCredito& = CantRegistros("cajabanc", CondicionSql$)
     If EsModoCredito& > 0 Then FORFASQL.chkModoCredito.Value = 1
     
     TIPOCOM% = 1                 ' valor por omision = FACTURA
     Select Case COMPROB$
           Case "FC": TIPOCOM% = TIPO_CBTE(COMPROB$, Variante$)
           Case "ND": TIPOCOM% = TIPO_CBTE(COMPROB$, Variante$)
           Case "NC": TIPOCOM% = TIPO_CBTE(COMPROB$, Variante$)
     End Select
     
     PVTANORMAL$ = CONTROL("FACTURA", "PUVENELE") 'P VTA. NORMAL
     PVTABCPTAL$ = CONTROL("FACTURA", "PUVENCAP") 'P.VTA. B.CAPITAL
     PVTAFAEXP$ = CONTROL("FACTURA", "PUVEFAEX")  'P.VTA. FACT.EXP.
     PVTMODOCREDITO$ = CONTROL("FACTURA", "PUVECRED")
     If EsModoCredito& > 0 Then FORFASQL.chkModoCredito.Value = 0
     '
     Select Case PVTA%
        Case XVALO(PVTANORMAL$), XVALO(PVTMODOCREDITO$)
            'Call SOLICITA_CAE_ID(IDENCAE$, CInt(EsModoCredito&))
           If EXISTE%(EJECTUTA$ + "wsfev1_cliente.exe") = 1 Then
              XXX = Shell(EJECTUTA$ & "wsfev1_cliente.exe AUTO " & IDENCAE$, 4)
              GoTo 99
           End If
           If EXISTE%(EJECTUTA$ + "wsfe_cliente.exe") = 1 Then
              XXX = Shell(EJECTUTA$ & "wsfe_cliente.exe AUTO " & IDENCAE$, 4)
           End If
           
        Case XVALO(PVTABCPTAL$)
          If EXISTE%(EJECTUTA$ + "wsbfev1_cliente.exe") = 1 Then
            XXX = Shell(EJECTUTA$ & "wsbfev1_cliente.exe AUTO " & IDENCAE$, 4)
          End If
          If EXISTE%(EJECTUTA$ + "wsbfe_cliente.exe") = 1 Then
            XXX = Shell(EJECTUTA$ & "wsbfe_cliente.exe AUTO " & IDENCAE$, 4)
          End If
          
        Case XVALO(PVTAFAEXP$)
          If EXISTE%(EJECTUTA$ + "wsfexv1_cliente.exe") = 1 Then
           XXX = Shell(EJECTUTA$ & "wsfexv1_cliente.exe AUTO " & IDENCAE$, 4)
          End If
          
        Case Else
           If MODOAUTO = "NO" Then Call COMUNI("No se Ha Encontrado el Punto de Venta " & PVTA%)
           Exit Sub
     End Select
     '
99   Timer1.Enabled = True    ' DISPARA LAZO DE ESPERA DE LA RESPUESTA DE AFIP
     '
End Sub


Private Sub Form_Load()
   UTILIZA_SKIN% = 1
   ModoRecuperacionDeComprobante = False
   Call APLICACIONSKINS(Me, Picture1)
   
   ModoSolicitudManual = False
End Sub



Private Sub mnuFacturaElectronicaComun_Click()

    Ruta$ = LUCOM$ & "FACELEC\"
    
    PuntoVenta% = XVALO(CONTROL("FACTURA", "PUVENELE")) 'P VTA. NORMAL
    
    If PuntoVenta% > 0 Then
        If EXISTE%(Ruta$ + "wsfev1_cliente.exe") = 0 Then
            Call COMUNI("Imposible Continuar.\ \Su Sistema No Cuenta con la Aplicación 'WSFEV1'.")
            Exit Sub
        End If
    
        RecuperacionPuntoVenta% = PuntoVenta%
        ReDim RecuperacionTipoComprobante%(6)
        RecuperacionTipoComprobante%(0) = 1 'FCA
        RecuperacionTipoComprobante%(1) = 2 'NDA
        RecuperacionTipoComprobante%(2) = 3 'NCA
        RecuperacionTipoComprobante%(3) = 6 'FCB
        RecuperacionTipoComprobante%(4) = 7 'NDB
        RecuperacionTipoComprobante%(5) = 8 'NCB
        RecuperacionModoCredito = False
        ReDim RecuperacionUltimoNumeroPorTipo&(6)
        
        tmrRecuperaUltimoNroCbte.Enabled = True
    Else
        Call COMUNI("Su Sistema No Esta Configurado Para Emitir Factura Electrónica.")
    End If
    
End Sub

Private Sub mnuInicioFacturaElectronica_Click()

    If CLACONTRA% > 0 Then
        OBX$ = OBJPLA$("ENTRAFECHA1", VDEF$)
        If OBX$ = "" Then Exit Sub
        '
        FE% = CVI(Mid$(OBX$, 1, 2))
        If FE% <= HOY(HO$) Then
            Call GRACONTROL("FACTURA", "ININUEFE", CStr(FE%))
        End If
    End If

End Sub

Private Sub mnuRecuperaCbteBienesDeCapital_Click()

    Ruta$ = LUCOM$ & "FACELEC\"
    
    PuntoVenta% = XVALO(CONTROL("FACTURA", "PUVENCAP")) 'P.VTA. B.CAPITAL
    
    If PuntoVenta% > 0 Then
        If EXISTE%(Ruta$ + "wsbfev1_cliente.exe") = 0 Then
            Call COMUNI("Imposible Continuar.\ \Su Sistema No Cuenta con la Aplicación 'WSBFEV1'.")
            Exit Sub
        End If
     
        RecuperacionPuntoVenta% = PuntoVenta%
        ReDim RecuperacionTipoComprobante%(6)
        RecuperacionTipoComprobante%(0) = 1 'FCA
        RecuperacionTipoComprobante%(1) = 2 'NDA
        RecuperacionTipoComprobante%(2) = 3 'NCA
        RecuperacionTipoComprobante%(3) = 6 'FCB
        RecuperacionTipoComprobante%(4) = 7 'NDB
        RecuperacionTipoComprobante%(5) = 8 'NCB
        RecuperacionModoCredito = False
        ReDim RecuperacionUltimoNumeroPorTipo&(6)
        
        tmrRecuperaUltimoNroCbte.Enabled = True
    Else
        Call COMUNI("Su Sistema No Esta Configurado Para Emitir Factura Electrónica Bientes de Capital.")
    End If

End Sub


Private Sub mnuRecuperaFacturaCredito_Click()

    Ruta$ = LUCOM$ & "FACELEC\"
    
    PuntoVenta% = XVALO(CONTROL("FACTURA", "PUVECRED"))  'P VTA. CREDITO
    If PuntoVenta% = 0 Then PuntoVenta% = XVALO(CONTROL("FACTURA", "PUVENELE")) 'P VTA. NORMAL
    
    If PuntoVenta% > 0 Then
        If EXISTE%(Ruta$ + "wsfev1_cliente.exe") = 0 Then
            Call COMUNI("Imposible Continuar.\ \Su Sistema No Cuenta con la Aplicación 'WSFEV2'.")
            Exit Sub
        End If
        
        RecuperacionPuntoVenta% = PuntoVenta%
        ReDim RecuperacionTipoComprobante%(3)
        RecuperacionTipoComprobante%(0) = 201 'FCA
        RecuperacionTipoComprobante%(1) = 202 'NDA
        RecuperacionTipoComprobante%(2) = 203 'NCA
        RecuperacionModoCredito = True
        ReDim RecuperacionUltimoNumeroPorTipo&(3)
        
        tmrRecuperaUltimoNroCbte.Enabled = True
    Else
        Call COMUNI("Su Sistema No Esta Configurado Para Emitir Factura de Crédito Electrónica.")
    End If
    
End Sub

Private Sub mnuRecuperaFacturaExpo_Click()

    Ruta$ = LUCOM$ & "FACELEC\"
    
    PuntoVenta% = XVALO(CONTROL("FACTURA", "PUVEFAEX")) 'P VTA. EXPORTACION
    
    If PuntoVenta% > 0 Then
        If EXISTE%(Ruta$ + "wsfexv1_cliente.exe") = 0 Then
            Call COMUNI("Imposible Continuar.\ \Su Sistema No Cuenta con la Aplicación 'WSFEXV1'.")
            Exit Sub
        End If
    
        RecuperacionPuntoVenta% = PuntoVenta%
        ReDim RecuperacionTipoComprobante%(3)
        RecuperacionTipoComprobante%(0) = 19 'FCE
        RecuperacionTipoComprobante%(1) = 20 'NDE
        RecuperacionTipoComprobante%(2) = 21 'NCE
        RecuperacionModoCredito = False
        ReDim RecuperacionUltimoNumeroPorTipo&(3)
        
        tmrRecuperaUltimoNroCbte.Enabled = True
    Else
        Call COMUNI("Su Sistema No Esta Configurado Para Emitir Factura Electrónica de Exportación.")
    End If
    
End Sub

'
Private Sub Timer1_Timer()
   '
   Static CONTA%
   CONTA% = CONTA% + 1 ' SI SUPERA LOS 40 SEGUNDOS DE DEMORA EN TRAER EL CAE PRESENTA MENSAJE
   If CONTA% > 60 Then
'      If COMALTER%("No Se Ha Podido Obtener el CAE \\Seguir Intentando\Cancelar Factura") <> 1 Then
         'Call FINAL("")
      Timer1.Enabled = False
''         If MODO% = 1 And ESTADO% = 1 Then Timer3.Enabled = True
''         If MODO% = 2 And ESTADO% = 1 Then ESTADO% = 0
'      End If
      CONTA% = 0
   End If
   '
'   IDENTCAE$ = TRIM$(Text5)
   KAE$ = LEECAE$    '  (IDENTCAE$)
   '
End Sub

Function LEECAE$() ' (IDENTCAE$)
    '
    CARPETOK$ = "C:\FLEXOFT\FACELEC\"
    LEECAE$ = ""
    Cae$ = ""
    TX1$ = ""
    IDENCAE$ = TRIM$(Text5)
    VtoCae$ = ""
    nufact$ = ""
    FechaEmision$ = ""
    '
    If EXISTE%(CARPETOK$ + "CAE_" + IDENCAE$ + ".TXT") > 0 Then
        NARCHIVO% = FILEOPEN%(CARPETOK$ + "CAE_" + IDENCAE$ + ".TXT", 1, 128)  'ABRE EL ARCHIVO DE LECTURA
        Do While Not EOF(NARCHIVO%) ' RECORRE
          Line Input #NARCHIVO%, XXXXX$ ' LEE CADA REGISTRO
          TX1$ = TX1$ + XXXXX$ + vbCrLf
          If InStr(UCase$(XXXXX$), UCase$("<cae>")) > 0 Then
             POS1% = InStr(UCase$(XXXXX$), UCase$("<cae>")) + 5
             TXT$ = Mid$(XXXXX$, POS1%)
             POS2% = InStr(TXT$, "</")
             If Cae$ = "" Then Cae$ = Mid$(TXT$, 1, POS2% - 1)
          ElseIf InStr(UCase$(XXXXX$), UCase$("<CodAutorizacion>")) > 0 Then
             POS1% = InStr(UCase$(XXXXX$), UCase$("<CodAutorizacion>")) + 17
             TXT$ = Mid$(XXXXX$, POS1%)
             POS2% = InStr(TXT$, "</")
             If Cae$ = "" Then Cae$ = Mid$(TXT$, 1, POS2% - 1)
             
             'caso de punto de venta 4
          ElseIf InStr(UCase$(XXXXX$), UCase$("<cbt_desde>")) > 0 Then
             POS1% = InStr(UCase$(XXXXX$), UCase$("<cbt_desde>")) + 11
             TXT$ = Mid$(XXXXX$, POS1%)
             POS2% = InStr(TXT$, "</")
             If nufact$ = "" Then nufact$ = Mid$(TXT$, 1, POS2% - 1)
          ElseIf InStr(UCase$(XXXXX$), UCase$("<CbteDesde>")) > 0 Then
             POS1% = InStr(UCase$(XXXXX$), UCase$("<CbteDesde>")) + 11
             TXT$ = Mid$(XXXXX$, POS1%)
             POS2% = InStr(TXT$, "</")
             If nufact$ = "" Then nufact$ = Mid$(TXT$, 1, POS2% - 1)
          ElseIf InStr(UCase$(XXXXX$), UCase$("<fecha_vto>")) > 0 Then
             POS1% = InStr(UCase$(XXXXX$), UCase$("<fecha_vto>")) + 11
             TXT$ = Mid$(XXXXX$, POS1%)
             POS2% = InStr(TXT$, "</")
             If VtoCae$ = "" Then VtoCae$ = Mid$(TXT$, 1, POS2% - 1)
          ElseIf InStr(UCase$(XXXXX$), UCase$("<CAEFchVto>")) > 0 Then
             POS1% = InStr(UCase$(XXXXX$), UCase$("<CAEFchVto>")) + 11
             TXT$ = Mid$(XXXXX$, POS1%)
             POS2% = InStr(TXT$, "</")
             If VtoCae$ = "" Then VtoCae$ = Mid$(TXT$, 1, POS2% - 1)
          ElseIf InStr(UCase$(XXXXX$), UCase$("<FchVto>")) > 0 Then
             POS1% = InStr(UCase$(XXXXX$), UCase$("<FchVto>")) + 8
             TXT$ = Mid$(XXXXX$, POS1%)
             POS2% = InStr(TXT$, "</")
             If VtoCae$ = "" Then VtoCae$ = Mid$(TXT$, 1, POS2% - 1)
             
             'caso de punto de venta 5
          ElseIf InStr(UCase$(XXXXX$), UCase$("<cbte_nro>")) > 0 Then
             POS1% = InStr(UCase$(XXXXX$), UCase$("<cbte_nro>")) + 10
             TXT$ = Mid$(XXXXX$, POS1%)
             POS2% = InStr(TXT$, "</")
             If nufact$ = "" Then nufact$ = Mid$(TXT$, 1, POS2% - 1)
          'FECHA DE VENCIMINTO DE CAE
          ElseIf InStr(UCase$(XXXXX$), UCase$("<fch_venc_cae>")) > 0 Then
             POS1% = InStr(UCase$(XXXXX$), UCase$("<fch_venc_cae>")) + 14
             TXT$ = Mid$(XXXXX$, POS1%)
             POS2% = InStr(TXT$, "</")
             If VtoCae$ = "" Then VtoCae$ = Mid$(TXT$, 1, POS2% - 1)
          ElseIf InStr(UCase$(XXXXX$), UCase$("<Fecha_cbte_orig>")) > 0 Then
             POS1% = InStr(UCase$(XXXXX$), UCase$("<Fecha_cbte_orig>")) + 17
             TXT$ = Mid$(XXXXX$, POS1%)
             POS2% = InStr(TXT$, "</")
             If FechaEmision$ = "" Then FechaEmision$ = Mid$(TXT$, 1, POS2% - 1)
          End If
        Loop
        Close NARCHIVO%
    End If
    '
    Text3 = TX1$
    '
    If XVALO(Cae$) > 0 Then
        Timer1.Enabled = False
        Text1 = Cae$
        Label6 = VtoCae$
        Text2 = nufact$
        lblFechaEmisionAfip = FechaEmision$
        ' If ESTADO% = 1 And MODO% > 0 Then ESTADO% = 0
    End If
    '
    If InStr(TRIM$(TX1$), "CANCELAR") > 0 Then
        Timer1.Enabled = False
        Call PROCESA_NO_OBTUVO_CAE
        'If MODO% = 1 And ESTADO% = 1 Then Timer3.Enabled = True
        'If MODO% = 2 And ESTADO% = 1 Then ESTADO% = 0
    End If
    '
    LEECAE$ = Cae$
    '
End Function


Sub PROCESA_NO_OBTUVO_CAE()
     '
     TX$ = UCase$(Text3)
     '
     If MODOAUTO = "NO" Then Call COMUNI("No se Pudo Obtener el Cae ")
     
'     If InStr(UCase$(TX$), "ID DE TRANSACCION NO REGISTRADO") > 0 Then
'           Call PIDENUEVOCAE
'     ElseIf InStr(UCase$(TX$), "NO ES CORRELATIVO AL ULTIMO") > 0 Then
'           Exit Sub
'           Call PIDENUEVOCAE
'     ElseIf InStr(UCase$(TX$), "CORRELATIVO") > 0 Then
'           Call OBTIENEIDEDEREEMPLAZO
'           Call PIDENUEVOCAE
'     ElseIf InStr(UCase$(TX$), "ID DE TRANSACCION NO EXISTE") > 0 Then
'           Call PIDENUEVOCAE
'     Else
'           If MODOAUTO = "NO" Then Call COMUNI("No se Pudo Obtener el Cae ")
'     End If

     '
''     Select Case XXX ... PROCESANDO LO QUE ESTA EN EL TEXT3
''        CASE ID NO REGISTRADO
''           Call PIDENUEVOCAE
''       CASE NO CORRELATIVO
''           Call OBTIENEIDEDEREEMPLAZO
''           CALL REEMPLAZA IDE EN CAJABANC
''
''       Case Else
''         mENSAJE Y EXIT SUB
''
''      End Select
End Sub
'
Sub PIDENUEVOCAE()
     '
 TXTRAZA$ = TXTRAZA$ & "PUNTO 55 = " & Now & vbCrLf
      
      MODO% = 2
      PVTA% = XVALO(Text4)
      IDENCAE$ = TRIM$(Text5)
      '
      PVTANORMAL$ = CONTROL("FACTURA", "PUVENELE") 'P VTA. NORMAL
      PVTABCPTAL$ = CONTROL("FACTURA", "PUVENCAP") 'P.VTA. B.CAPITAL
      PVTAFAEXP$ = CONTROL("FACTURA", "PUVEFAEX")  'P.VTA. FACT.EXP.
      PVTMODOCREDITO$ = CONTROL("FACTURA", "PUVECRED")
      '
TXTRAZA$ = TXTRAZA$ & "PUNTO 56 = " & Now & vbCrLf
    
      If PVTA% < 1 And IDENCAE$ = "" Then
          If MODOAUTO = "NO" Then Call COMUNI("Faltan Datos ")
         Exit Sub
      End If
      '
      Text1 = ""
      Text2 = ""
      Text3 = ""
      Label6 = ""
      '
      Select Case PVTA%
         '
         Case XVALO(PVTANORMAL$)
              Call SOLICITA_CAE_ID(IDENCAE$)
         Case XVALO(PVTABCPTAL$)
              Call SOLICITA_CAE_BCAP_ID(IDENCAE$)
         Case XVALO(PVTAFAEXP$)
              Call SOLICITA_CAE_EXPO_ID(IDENCAE$)
         Case XVALO(PVTMODOCREDITO$)
              Call SOLICITA_CAE_ID(IDENCAE$, 1)
      End Select
TXTRAZA$ = TXTRAZA$ & "PUNTO 63 = " & Now & vbCrLf
      Timer1.Enabled = True
      '
End Sub

Sub OBTIENEIDEDEREEMPLAZO()
   '
       IDENCAE$ = TRIM$(Text5)
       '
       'BUSCO EL VALOR MAXIMO DEL IDENCAE (SE UTILIZA UN CAST POR QUE ES UN CAMPO DE TEXTO)
'       IdentiCaeMax$ = TRIM$(Str$(1 + XVALO(VALOBADA("CAJABANC", "MAX(cast(substring(IDCAE,1,LEN(IDCAE))as INT))", " TIPOMOVIM = 'FVEN'", "NOMESS"))))
       '
        IDCAENUX& = XVALO(VALOBADA("CAJABANC", "MAX(IDCAENUM)", " TIPOMOVIM = 'FVEN'", "NOMESS"))
        If IDCAENUX& < 1 Then
            IdentiCaeMax$ = TRIM$(Str$(1 + XVALO(VALOBADA("CAJABANC", "MAX(cast(substring(IDCAE,1,LEN(IDCAE))as INT))", " TIPOMOVIM = 'FVEN'", "NOMESS"))))
        Else
            IdentiCaeMax$ = SAFETEXT$(1 + IDCAENUX&)
        End If
        '
       'GUARDO EL CAE NUEVO REEMPLAZANDO EL ANTERIOR
       'CONDI$ = "IDCAE = '" & IDENCAE$ & "'"
       Condi$ = "IDCAENUM = " & XVALO(IDENCAE$)
       SQLX$ = " SELECT IDCAE, IDCAENUM FROM CAJABANC "
       SQLX$ = SQLX$ + " WHERE IDCAE = '" & IDENCAE$ & "'"
       Dim RSCBANC As ADODB.Recordset
       Set RSCBANC = New ADODB.Recordset
25     On Error Resume Next
       RSCBANC.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
            On Error GoTo 0
            Call CapturaErrorOpen
            GoTo 25
       End If
       On Error GoTo 0
       With RSCBANC
          Do While Not .EOF
             !IDCAENUM = XVALO(IdentiCaeMax$)
             !IDCAE = IdentiCaeMax$
             .Update
             .MoveNext
          Loop
       End With
       RSCBANC.Close
       Set RSCBANC = Nothing
   
'      Call ACTUREGISTRO("CAJABANC", "IDCAENUM", CONDI$, XVALO(IdentiCaeMax$), REAF&, "NOMESS")
'      Call ACTUREGISTRO("CAJABANC", "IDCAE", CONDI$, IdentiCaeMax$, REAF&, "NOMESS")
   '
   '   ACTUALIZO LA CAJA DE TEXTO DEL IDENTIFCADOR DE CAE
       Text5 = IdentiCaeMax$
   '
End Sub

Private Sub Command16_Click()
    '
    ' ESTE BOTÓN ES PARA RE-IMPRIMIR COMPROBANTES
    '
    Condi$ = "TipoMovim = 'FVEN' AND NUMECOMP > 0 ORDER BY FECHEMISI"
    Call CARGALISEL("!LISCOMP", "CAJABANC+DEUDOR12/CAJABANC.NUCLIEN = DEUDOR12.NUME_CLI", "CAJABANC.CODICOM_LAR/A18;DEUDOR12.NUME_CLI/F6;DEUDOR12.RASO_CLI/A18;FechEmisi/D8", Condi$, "DISTINCT")
    '
10  Call SELECHO("!LISCOMP/Lista de Comprobantes de Venta")
    NUFACTU$ = TRIM$(VALACT1("!LISCOMP"))
    If NUFACTU$ = "" Then Exit Sub
    
    TXT$ = VALACT2("!LISCOMP")
    NN = InStr(1, TXT$, " ")
    NCLI = Val(Mid(TXT$, 1, NN - 1))
    
    LETR$ = Mid$(NUFACTU$, 18, 1)
        
    Call PRINFACTU(NUFACTU$, NCLI, LETR$, "PRINT")
    GoTo 10
    
End Sub

Sub Command22_Click()
   
   Condi$ = "NUMECOMP < 0  AND TIPOMOVIM = 'FVEN'"
   INDICE& = CantRegistros("CAJABANC", Condi$, "NOMESS")
   If INDICE& > 0 Then ReDim FACTU$(INDICE&)
   '
   i& = 0
   Call ABRECABAN
   SQLX$ = "SELECT IDCAE FROM CAJABANC WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE " & Condi$
   SQLX$ = SQLX$ + " ORDER BY FEHOREAL "
   Dim CAJATMP As New ADODB.Recordset
   Set CAJATMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   With CAJATMP
      '
      Do While Not .EOF
         DOCU$ = SAFETEXT$(!IDCAE)
         If TRIM$(DOCU$) <> "" Then
            i& = i& + 1
            FACTU$(i&) = DOCU$
         End If
      .MoveNext
      Loop
      
   End With
   CAJATMP.Close
   Set CAJATMP = Nothing
   Timer2.Enabled = True
   ESTADO% = 0
   CONTAD& = 0
   
End Sub


Private Sub Command3_Click()

   Command3.Enabled = False
   '
   Condi$ = "TipoMovim = 'FVEN' AND NUMECOMP < 0 "
   Condi$ = Condi$ + " AND IDCAE <> '' ORDER BY FECHEMISI"
   Call CARGALISEL("FACTUSEL", "CAJABANC+DEUDOR12/CAJABANC.NUCLIEN = DEUDOR12.NUME_CLI", "CAJABANC.IDCAE/A14;DEUDOR12.NUME_CLI/F6;DEUDOR12.RASO_CLI/A18;FechEmisi/D8", Condi$, "DISTINCT")
   '
   Call SELECHO("FACTUSEL/Facturas Pendientes de Solicitud de Cae")             'Muestra Docu de Credito
   IDCAEX$ = TRIM$(VALACT1("FACTUSEL"))
   If IDCAEX$ <> "" Then
     Text5 = IDCAEX$
   End If
   '
   Command3.Enabled = True
   
'   Command3.Enabled = False
'   '
'   Condi$ = "TipoMovim = 'FVEN' AND NUMECOMP < 0 AND modoopcimovim<>'cr'"
'
'   ' SI ESTA CONFIGURADA LA FECHA DE INICIO DE LA NUEVA VERSION DE FACTURA ELECTRONICA COMUN
'   ' ACA SOLO PRESENTA LAS ANTERIORES A DICHA FECHA. LAS POSTERIORES SE RECUPERAN DESDE LA
'   ' OPCION HABILITADA EN EL MENU
'   PtoVtaFacturaElectronicaComun& = XVALO(Control("FACTURA", "PUVENELE")) 'P VTA. NORMAL
'   PtoVtaBienesDeCapital& = XVALO(Control("FACTURA", "PUVENCAP"))
'   InicioNuevaFacturaElectronicaComun% = XVALO(Control("FACTURA", "ININUEFE"))
'   If InicioNuevaFacturaElectronicaComun% > 0 Then
'        Condi$ = Condi$ & " and (numetalo<>" & CStr(PtoVtaBienesDeCapital&) & ") and (numetalo<>" & CStr(PtoVtaFacturaElectronicaComun&) _
'        & " or (numetalo=" & CStr(PtoVtaFacturaElectronicaComun&) & " and fechconta<'" & FETEXCOR1$(InicioNuevaFacturaElectronicaComun%) & "'))"
'   End If
'
'   Condi$ = Condi$ + " AND IDCAE <> '' ORDER BY FECHEMISI"
'   Call CARGALISEL("FACTUSEL", "CAJABANC+DEUDOR12/CAJABANC.NUCLIEN = DEUDOR12.NUME_CLI", "CAJABANC.IDCAE/A14;DEUDOR12.NUME_CLI/F6;DEUDOR12.RASO_CLI/A18;FechEmisi/D8", Condi$, "DISTINCT")
'   '
'   Call SELECHO("FACTUSEL/Facturas Pendientes de Solicitud de Cae")             'Muestra Docu de Credito
'   IDCAEX$ = TRIM$(VALACT1("FACTUSEL"))
'   If IDCAEX$ <> "" Then
'     Text5 = IDCAEX$
'   End If
'   '
'   Command3.Enabled = True
End Sub

''Sub Command5_Click()
''     '
''      Command5.Enabled = False
''      '
''      MODO% = 1
''      PVTA% = XVALO(Text4)
''      IDENCAE$ = TRIM$(Text5)
''      If PVTA% > 0 And IDENCAE$ <> "" Then
''        Call PIDECAE(PVTA%, IDENCAE$)
''      Else
''        If MODOAUTO = "NO" Then Call COMUNI("Faltan Datos")
''      End If
''      '
''      Command5.Enabled = True
''      '
''End Sub

Function TIPO_DOC$(NC%)
    'TIPO DE DOCUMENTO
    'NC% = XVALO(FORFACPRINT.Text1)
    POSDEIVA% = PIVACLI%(NC%)
    Select Case POSDEIVA%
        Case 0: valor$ = 96
        Case 1: valor$ = 80
        Case 2: valor$ = 80
        Case 3: valor$ = 80
        Case 4: valor$ = 80
        Case 5: valor$ = 80
        Case 6: valor$ = 80
        Case 7: valor$ = 80
    End Select
    TIPO_DOC$ = valor$
End Function


Function NRO_DOC$(NC%)
    NRO_DOC$ = ""
    POXIVA% = PIVACLI%(NC%)
    '
    NUMERO_CUIT_DOC1$ = CuitCli$(NC%)
    NUMERO_CUIT_DOC$ = REPLACAR(NUMERO_CUIT_DOC1$, "-", "")
    '
    If POXIVA% < 1 And TRIM$(NUMERO_CUIT_DOC$) = "" Then NUMERO_CUIT_DOC$ = "0"
    '
    If Not IsNumeric(NUMERO_CUIT_DOC$) Then
      If MODOAUTO = "NO" Then Call MENSERR(24, "Documento Nro.: " & NUMERO_CUIT_DOC$ & " No Válido\El Dato debe ser Numérico\Modifique y Vuelva a Cargar la Factura ")
      Call FINAL("")
    End If
    '
    NRO_DOC$ = NUMERO_CUIT_DOC$
    '
End Function


Function IMP_TOT_CONC$(PERCEPIVA#, PERCEPIB#, IMPNONETO#)

   IMP_TOT_CONC$ = "" ' SUMA DE IVA NO GRAVADO + PERCEPCION DE IVA + PERCEPCION DE INGRESOS BRUTOS
   IMPTOTACONC$ = FORMATNUM$(XVALO(IMPNONETO#) + XVALO(PERCEPIVA#) + XVALO(PERCEPIB#), "F12.2")
   If TRIM$(IMPTOTACONC$) = "" Then IMPTOTACONC$ = "0"
   IMP_TOT_CONC$ = TRIM$(IMPTOTACONC$)
   
End Function

Function IMPTO_LIQ$()
  IMPTO_LIQ$ = ""
  IMPTOLIQUID$ = FORMATNUM$(XVALO(FORFACPRINT.Total(3)) + XVALO(FORFACPRINT.Total(4)) + XVALO(FORFACPRINT.Total(5)), "F12.2")
  IMPTO_LIQ$ = TRIM$(IMPTOLIQUID$)
End Function

Function MONEDA_CAE$()
    '
    MONE$ = ""
    If MONEMI% = 1 Then
         MONE$ = "PES"
      ElseIf MONEMI% = 2 Then
         MONE$ = "DOL"
      Else
         DEMONE$ = UCase(DEMONECO$(MONEMI%))
         If Left$(DEMONE$, 4) = "EURO" Then
              MONE$ = "060"
           ElseIf Left$(DEMONE$, 4) = "REAL" Then
              MONE$ = "012"
         End If
    End If
    '
99  MONEDA_CAE$ = MONE$
    '
End Function

Function TIPO_CBTE$(PREFI$, LETRA$)
    
    TIPOCOMP$ = UCase$(PREFI$ & LETRA$)
    Select Case TIPOCOMP$
        '
        Case "FCA": valor$ = "1": If FORFASQL.chkModoCredito.Value = 1 Then valor$ = "201"  'FACTURA A
        Case "NDA": valor$ = "2": If FORFASQL.chkModoCredito.Value = 1 Then valor$ = "202"  'N DEBITO A
        Case "NCA": valor$ = "3": If FORFASQL.chkModoCredito.Value = 1 Then valor$ = "203" 'N CREDITO A
        Case "FCB": valor$ = "6": If FORFASQL.chkModoCredito.Value = 1 Then valor$ = "206" 'FACTURA B
        Case "NDB": valor$ = "7": If FORFASQL.chkModoCredito.Value = 1 Then valor$ = "207" 'N DEBITO B
        Case "NCB": valor$ = "8": If FORFASQL.chkModoCredito.Value = 1 Then valor$ = "208" 'N CREDITO B
        Case "FCE": valor$ = "19"
        Case "NDE": valor$ = "20"
        Case "NCE": valor$ = "21"
        '
    End Select
    '
    TIPO_CBTE$ = valor$
    '
End Function

'Private Sub Command6_Click()
'
'      SOLOCAE% = 1
'      '
'      Call Command5_Click
'
'End Sub

Private Sub Text1_GotFocus()
On Error Resume Next
Text5.SetFocus
On Error GoTo 0
End Sub

Private Sub Text2_Change()

   If ModoRecuperacionDeComprobante = True Then Exit Sub
   
   If XVALO(Text1) > 0 Then
      PREFI$ = "0000" & TRIM$(Text4): PREFI$ = Right(PREFI$, 4)
      NumeroFactura$ = "00000000" & TRIM$(Text2): NumeroFactura$ = Right(NumeroFactura$, 8)
      NUMFACTUNUE$ = Mid$(Text6, 1, 3) & PREFI$ + "-" + NumeroFactura & Mid$(Text6, 17, 2)
      NUFACTUVIE$ = TRIM$(Text6)
      
      If XVALO(Text5) > 0 And NUFACTUVIE$ <> "" Then
        Call GRABANUEFACTU(NUFACTUVIE$, NUMFACTUNUE$)
      Else
        Call COMUNI("Imposible Continuar.\Numero de Id Erroneo o Falta de Numero Provisorio.\\Favor Comunicarse Con Mesa de Ayuda Flexoft")
      End If
   End If
End Sub


Private Sub Text2_gotfocus()
On Error Resume Next
Text5.SetFocus
On Error GoTo 0
End Sub



Private Sub Text5_Change()
    Call ABRECABAN
 TXTRAZA$ = TXTRAZA$ & "PUNTO 53 = " & Now & vbCrLf
    DOCUM$ = TRIM$(Text5)
    'IDENDIFICADOR DE CAE
    
'    CONDI$ = "IDCAE = '" & DOCUM$ & "'"
    'SE MODIFICA PARA QUE BUSQUE POR EL IDCAENUM QUE  ES MAS RAPIDO Y ESTA INDEXADO
    Condi$ = "IDCAENUM = " & XVALO(DOCUM$) & " AND TIPOMOVIM='FVEN'"
    'VER SI SE PUEDE AGREGAR LA CONDICION QUE SEA TIPOMOVIM  = 'FVEN'
    Cae$ = TRIM$(VALOBADA("CAJABANC", "CAE", Condi$))
    'EN CASO QUE NO DEVUELVA NADA UTILIZA LA CONDICION ANTERIOR
'    If CAE$ = "" Then
'      CONDI$ = "IDCAE = '" & DOCUM$ & "'"
'      CAE$ = TRIM$(VALOBADA("CAJABANC", "CAE", CONDI$))
'    End If
 TXTRAZA$ = TXTRAZA$ & "PUNTO 53 = " & Now & vbCrLf
    If Cae$ <> "" Then
        Call COMUNI("Este Comprobante ya posee CAE")
        Text6 = "": Exit Sub
    End If
    
    Text6 = VALOBADA("CAJABANC", "CODICOM_LAR", Condi$)
 TXTRAZA$ = TXTRAZA$ & "PUNTO 54 = " & Now & vbCrLf
    '
    'PUNTO DE VENTA
    TIP$ = REPLACAR$(Mid$(Text6, 4, 4), "X", "0")
    Text4 = XVALO(TIP$)
    
    'limpio el resto de los casilleros
    Text1 = ""
    Text2 = ""
    Text3 = ""
    Label6 = ""

End Sub

Private Sub Text6_GotFocus()
On Error Resume Next
Text5.SetFocus
On Error GoTo 0
End Sub

Sub GRABANUEFACTU(NUFACTUVIE$, NUFACTUNUE$)
   Call ABRECABAN
   '
TXTRAZA$ = TXTRAZA$ & "PUNTO 66 = " & Now & vbCrLf
   
   TICOMX$ = Left(NUFACTUVIE$, 2)
   NUMECOMPROB$ = "000000" & XVALO(Text2): NUMECOMPROB$ = Right(NUMECOMPROB$, 6)
   LETRACOMP$ = Right(TRIM$(NUFACTUVIE$), 1)
   FACTUCOR$ = TICOMX$ + "-" + Right(NUMECOMPROB$, 6) + "-" + LETRACOMP$
   Cae$ = TRIM$(Text1)
   '
TXTRAZA$ = TXTRAZA$ & "PUNTO 67 = " & Now & vbCrLf
   
   Condi$ = "CODICOM_LAR = '" & NUFACTUVIE$ & "'"
   NCLI = XVALO(VALOBADA("CAJABANC", "NUCLIEN", Condi$, "NOMESS"))
   FF% = HOY(HO$) ': FECHA$ = FETEXCOR1$(FF%)
   FechaEmisionAfip% = FECHANUM(lblFechaEmisionAfip)
   If FechaEmisionAfip% > 0 Then FF% = FechaEmisionAfip%
   FF1% = FF%
   FECHVENCICAE$ = TRIM$(Label6)
   
   If REDOCEL.MODORECU > 0 Then FF1% = FECHANUM("31/12/10")
   '
TXTRAZA$ = TXTRAZA$ & "PUNTO 68 = " & Now & vbCrLf
   
   COBA$ = ARMACODIGOBARRA(NUFACTUNUE$, Cae$, FECHVENCICAE$)
TXTRAZA$ = TXTRAZA$ & "PUNTO 69 = " & Now & vbCrLf
   '
   CodigoQRAfip$ = ARMACODIGOQR(Text5)
   '
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + " WHERE CODICOM_LAR = '" & NUFACTUVIE$ & "'"
    
    Dim RSCBANC As ADODB.Recordset
    Set RSCBANC = New ADODB.Recordset
25  On Error Resume Next
TXTRAZA$ = TXTRAZA$ & "PUNTO 70 = " & Now & vbCrLf
    RSCBANC.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
TXTRAZA$ = TXTRAZA$ & "PUNTO 71 = " & Now & vbCrLf
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
        On Error GoTo 0
        Call CapturaErrorOpen
        GoTo 25
    End If
    On Error GoTo 0
    With RSCBANC
      Do While Not .EOF
       !CodiCom_Cor = FACTUCOR$
       !NumeComp = XVALO(NUMECOMPROB$)
       !Cae = Cae$
       !fechconta = CVDAT(FF1%) 'FECHA1$
       !FECHEMISI = CVDAT(FF%)  'FECHA$
       !FechVtoCae = FECHVENCICAE$
       !COD_BARRA = COBA$
       !CodigoQRAfip = CodigoQRAfip$
       !codicom_lar = NUFACTUNUE$
       .Update
       .MoveNext
      Loop
    End With
    RSCBANC.Close
    Set RSCBANC = Nothing
    '
TXTRAZA$ = TXTRAZA$ & "PUNTO 72 = " & Now & vbCrLf
  
    '''''****************
    SQLX$ = " SELECT CODICOM_COR, FECHEMISI, CODICOM_LAR FROM SADEUPEN "
    SQLX$ = SQLX$ + " WHERE CODICOM_LAR = '" & NUFACTUVIE$ & "'"
    Dim RSSDP As ADODB.Recordset
    Set RSSDP = New ADODB.Recordset
45  On Error Resume Next
    RSSDP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
        On Error GoTo 0
        Call CapturaErrorOpen
        GoTo 45
    End If
    On Error GoTo 0
    With RSSDP
      Do While Not .EOF
       !CodiCom_Cor = FACTUCOR$
       !FECHEMISI = CVDAT(FF%)
       !codicom_lar = NUFACTUNUE$
       .Update
       .MoveNext
      Loop
    End With
    RSSDP.Close
    Set RSSDP = Nothing
   
TXTRAZA$ = TXTRAZA$ & "PUNTO 73 = " & Now & vbCrLf
    ''''''''''*********
    If TICOMX$ = "NC" Then
        Condi$ = "codocap_lar='" & NUFACTUVIE$ & "'"
        Call ACTUREGISTRO("aplicobra", "codocap_cor", Condi$, FACTUCOR$, REGAF&, "NOMESS")
        Call ACTUREGISTRO("aplicobra", "codocap_lar", Condi$, NUFACTUNUE$, REGAF&, "NOMESS")
    End If
    '
   If REDOCEL.MODORECU > 0 Then
     Condi$ = "CODICOM_LAR = '" & NUFACTUVIE$ & "'"
     Call ACTUREGISTRO("APLICOBRA", "CODOCOR_COR", Condi$, FACTUCOR$, REGAF&, "NOMESS")
     Call ACTUREGISTRO("APLICOBRA", "CODOCOR_LAR", Condi$, NUFACTUNUE$, REGAF&, "NOMESS")
   End If
   '
   Call ABRESTOCKS
   '
TXTRAZA$ = TXTRAZA$ & "PUNTO 74 = " & Now & vbCrLf
   
   TABLA1$ = "MOVISTO"
   'VALIDO SI ES UNA REPARACION
   Condi$ = "DOSECLAR = '" & NUFACTUVIE$ & "'"
   CANTREGREPA& = XVALO(CantRegistros("MOVIREPA", Condi$, "NOMESS"))
   If CANTREGREPA& > 0 Then TABLA1$ = "MOVIREPA"
   '
   SQLX$ = "SELECT * FROM " & TABLA1$
   SQLX$ = SQLX$ + " WHERE DOSECLAR = '" & NUFACTUVIE$ & "'"
   
   Dim RSMVSTO As ADODB.Recordset
   Set RSMVSTO = New ADODB.Recordset
26 On Error Resume Next
TXTRAZA$ = TXTRAZA$ & "PUNTO 75 = " & Now & vbCrLf
   RSMVSTO.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
TXTRAZA$ = TXTRAZA$ & "PUNTO 75 = " & Now & vbCrLf
   
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
        On Error GoTo 0
        Call CapturaErrorOpen
        GoTo 26
   End If
   On Error GoTo 0
   '
   With RSMVSTO
     Do While Not .EOF
       !DoSecCor = FACTUCOR$
       !DOSECLAR = NUFACTUNUE$
       .Update
       .MoveNext
     Loop
   End With
   RSMVSTO.Close
   Set RSMVSTO = Nothing
 TXTRAZA$ = TXTRAZA$ & "PUNTO 76 = " & Now & vbCrLf
  '
   SQLX$ = "SELECT * FROM " & TABLA1$
   SQLX$ = SQLX$ + " WHERE DOPRILAR = '" & NUFACTUVIE$ & "'"
   '
   ' ESTO ES PARA REEMPLAZAR EL REMITO SI QUEDO COMO FACTURA SIN REMITO
   Dim RSMVSTO1 As ADODB.Recordset
   Set RSMVSTO1 = New ADODB.Recordset
27 On Error Resume Next
TXTRAZA$ = TXTRAZA$ & "PUNTO 77 = " & Now & vbCrLf
   RSMVSTO1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
TXTRAZA$ = TXTRAZA$ & "PUNTO 78 = " & Now & vbCrLf
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
        On Error GoTo 0
        Call CapturaErrorOpen
        GoTo 27
    End If
    On Error GoTo 0
    With RSMVSTO1
      Do While Not .EOF
       !DoPriCor = FACTUCOR$
       !doprilar = NUFACTUNUE$
       .Update
       .MoveNext
      Loop
    End With
    RSMVSTO1.Close
    Set RSMVSTO1 = Nothing


'   Call ACTUREGISTRO(TABLA1$, "DOSECCOR", CONDI$, FACTUCOR$, REGAF&, "NOMESS")
'   Call ACTUREGISTRO(TABLA1$, "DOSECLAR", CONDI$, NUFACTUNUE$, REGAF&, "NOMESS")
   '
'   ' ESTO ES PARA REEMPLAZAR EL REMITO SI QUEDO COMO FACTURA SIN REMITO
'   CONDI2$ = "DOPRILAR = '" & NUFACTUVIE$ & "'"
'   If CantRegistros(TABLA1$, CONDI2$) > 0 Then
'      Call ACTUREGISTRO(TABLA1$, "DOPRICOR", CONDI2$, FACTUCOR$, REGAF&, "NOMESS")
'      Call ACTUREGISTRO(TABLA1$, "DOPRILAR", CONDI2$, NUFACTUNUE$, REGAF&, "NOMESS")
'   End If

   '
   Call CREATABLA_EQUIPOS
    SQLX$ = " SELECT NUMFACTURA FROM TANUMSE "
    SQLX$ = SQLX$ + " WHERE NUMFACTURA = '" & NUFACTUVIE$ & "'"
    Dim TSTNSE As ADODB.Recordset
    Set TSTNSE = New ADODB.Recordset
55  On Error Resume Next
TXTRAZA$ = TXTRAZA$ & "PUNTO 79 = " & Now & vbCrLf
    TSTNSE.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
TXTRAZA$ = TXTRAZA$ & "PUNTO 80 = " & Now & vbCrLf
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
        On Error GoTo 0
        Call CapturaErrorOpen
        GoTo 55
    End If
    On Error GoTo 0
    With TSTNSE
      Do While Not .EOF
       !NumFACTURA = NUFACTUNUE$
       .Update
       .MoveNext
      Loop
    End With
    TSTNSE.Close
    Set TSTNSE = Nothing

'   CONDI$ = "NUMFACTURA = '" & NUFACTUVIE$ & "'"
'   Call ACTUREGISTRO("TANUMSE", "NUMFACTURA", CONDI$, NUFACTUNUE$, REGAF&, "NOMESS")
   '
   'PONGO EN 0 LA BANDERA DE MARCA DE MODO DE RECUPERACION
   REDOCEL.MODORECU = 0
TXTRAZA$ = TXTRAZA$ & "PUNTO 81 = " & Now & vbCrLf
   '
   ' SI EXISTE EN REPARA ACTUALIZA  NUMERO DE FACTURA
   If InStr(EMPREAC$, "MEDITEA") > 0 Then
     Condi$ = "FACTURA = '" & NUFACTUVIE$ & "'"
     If CantRegistros&("REPARA", Condi$, "MOMESS") > 0 Then
        NUREP& = XVALO(VALOBADA("REPARA", "NUMREPA", Condi$, "NOMESS"))
        NUMESER$ = VALOBADA("REPARA", "NUMEROSERIE", Condi$, "NOMESS")
        RACLI$ = rasocli$(NCLI)
        'ACTUALIZO EN REPARA EL NUMERO DE FACTURA
        If Left$(NUFACTUVIE$, 2) = "FC" Then
          Call ACTUREGISTRO("REPARA", "FACTURA", Condi$, NUFACTUNUE$, REGAF&, "NOMESS")
        Else
          Call ACTUREGISTRO("REPARA", "IMPORTE", Condi$, 0, REGAF&, "NOMESS")
          Call ACTUREGISTRO("REPARA", "FACTURA", Condi$, "", REGAF&, "NOMESS")
        End If
        '
        'GRABA EN ANOTADOR DE NUMERO DE SERIE
        CUERPO$ = "Número de Reparación: " & NUREP& & vbCrLf
        CUERPO$ = CUERPO$ + "Comprobante: " & NUFACTUNUE$ & vbCrLf
        CUERPO$ = CUERPO$ + "Cliente: " & RACLI$ + " (" & NCLI & ")"
        Call ANOTANUSE(NUMESER$, CUERPO$, , "NOMESS")
     End If
   End If
TXTRAZA$ = TXTRAZA$ & "PUNTO 82 = " & Now & vbCrLf
   '
   If TABLA1$ = "MOVISTO" Then
        PEDIANT& = 0
        SQLX$ = "SELECT * FROM MOVISTO WITH(NOLOCK) "
        SQLX$ = SQLX$ + " WHERE DOSECLAR = '" & NUFACTUNUE$ & "' ORDER BY NUPEDCLI "
        Set RSMVSTO = READSET(SQLX$)
TXTRAZA$ = TXTRAZA$ & "PUNTO 83 = " & Now & vbCrLf
        With RSMVSTO
          Do While Not .EOF
                CI% = XVALO(!COD_INTE)
                NPEDIDX& = XVALO(!NUPEDCLI)
                DOPRICO$ = SAFETEXT$(!DoPriCor)
                If NPEDIDX& <> PEDIANT& Then
                   If PEDIANT& <> 0 Then
                       Call AGREGA_ANOTAPED(PEDIANT&, Space$(10) + String$(56, "-"))
                       Call AGREGA_ANOTAPED(PEDIANT&, String$(88, "="))
                   End If
                   PEDIANT& = NPEDIDX&
                   TTXX$ = ""
                   TTXX$ = "Remito " + TRIM$(Mid$(DOPRICO$, 4))
                   TTXX$ = TTXX$ + " - Factura " + TRIM$(NUFACTUNUE$)
                   TTXX$ = TTXX$ + ":"
                   '
                   TTXYZ$ = Space$(76)
                   HOII% = HOY(HOS$)
                   Call REPLA(TTXYZ$, HOS$, 1, 8)
                   Call REPLA(TTXYZ$, TTXX$, 11, 56)
                   Call REPLA(TTXYZ$, USERTER$, 69, 24)
                   Call AGREGA_ANOTAPED(NPEDIDX&, TTXYZ$)
                   Call REPLA(TTXYZ$, Space$(88), 1, 88)
                   Call AGREGA_ANOTAPED(NPEDIDX&, TTXYZ$)
                   '
                End If
                .MoveNext
          Loop
        End With
        On Error Resume Next
        RSMVSTO.Close
        Set RSMVSTO = Nothing
        On Error GoTo 0
TXTRAZA$ = TXTRAZA$ & "PUNTO 83 = " & Now & vbCrLf
        
        If PEDIANT& > 0 Then
            Call AGREGA_ANOTAPED(PEDIANT&, Space$(10) + String$(56, "-"))
            Call AGREGA_ANOTAPED(PEDIANT&, String$(88, "="))
        End If
   End If
TXTRAZA$ = TXTRAZA$ & "PUNTO 84 = " & Now & vbCrLf
'
'ACTUALIZO LA TABLA DE ANTIICPOS - SI NO EXSISTE NO ENTRA
    SQLX$ = "SELECT * FROM FACTUANTI "
    SQLX$ = SQLX$ + " WHERE DOCU_ORIG = '" & NUFACTUVIE$ & "'"
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
35  On Error Resume Next
    rs.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
        On Error GoTo 0
        Call CapturaErrorOpen
        GoTo 35
    End If
    On Error GoTo 0
    With rs
      Do While Not .EOF
       !NumeCom = XVALO(NUMECOMPROB$)
       !DOCU_ORIG = NUFACTUNUE$
       .Update
       .MoveNext
      Loop
    End With
    rs.Close
    Set rs = Nothing
    '
    'PARA QUE ACTUALICE SI ES APLICACION DE FACTURA
    SQLX$ = "SELECT * FROM FACTUANTI "
    SQLX$ = SQLX$ + " WHERE DOCU_APLI  = '" & NUFACTUVIE$ & "'"
    Dim rs1 As ADODB.Recordset
    Set rs1 = New ADODB.Recordset
65  On Error Resume Next
    rs1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
        On Error GoTo 0
        Call CapturaErrorOpen
        GoTo 65
    End If
    On Error GoTo 0
    With rs1
      Do While Not .EOF
       !NumeCom = XVALO(NUMECOMPROB$)
       !DOCU_APLI = NUFACTUNUE$
       .Update
       .MoveNext
      Loop
    End With
    rs1.Close
    Set rs1 = Nothing
'FIN FACTURA ANTICIPO
'

    SQLX$ = "SELECT * FROM HojaTrabajoDifCambio "
    SQLX$ = SQLX$ + " WHERE NCLIEN = " & NCLI & " AND  NOTADEBITOGENERADA  = '" & NUFACTUVIE$ & "'"
    Dim RS2 As ADODB.Recordset
    Set RS2 = New ADODB.Recordset
70  On Error Resume Next
    RS2.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
        On Error GoTo 0
        Call CapturaErrorOpen
        GoTo 70
    End If
    On Error GoTo 0
    With RS2
      Do While Not .EOF
       !NOTADEBITOGENERADA = NUFACTUNUE$
       !NumeComp = XVALO(NUMECOMPROB$)
       .Update
       .MoveNext
      Loop
    End With
    RS2.Close
    Set RS2 = Nothing
    '
    'ITEMS SIN CARGO MEDITEA ACTUALIZA LA REFERENCIA DEL ASISENTO CON LA FACTURA ELECTRONICA
    If CONTROL("ADMIPED", "ITEMSCGO") = "SI" Then
        If TRIM$(NUFACTUVIE$) <> "" Then
            Condi$ = "Refeasien LIKE '%" & NUFACTUVIE$ & "%'"
            SQLX$ = "Update ENCABEASIEN "
            SQLX$ = SQLX$ + " Set Refeasien = replace(Refeasien,"
            SQLX$ = SQLX$ + "'" & NUFACTUVIE$ & "'"
            SQLX$ = SQLX$ + ","
            SQLX$ = SQLX$ + "'" & NUFACTUNUE$ & "'"
            SQLX$ = SQLX$ + ")"
            SQLX$ = SQLX$ + " WHERE " & Condi$
            FLEXCONN.Execute (SQLX$)
            'Update ENCABEASIEN  Set Refeasien = replace(Refeasien,'FC-XXX1-00001234','FC-0001-00004567') WHERE Refeasien LIKE '%FC-XXX1-00001234%'
            'cadena 1 = es el campo donde esta lo que queremos reemplazar
            'cadena 2 = es el string que queremos reemplazar
            'cadena 3 = es el string por el que vamos a reemplazar el anterior
        End If
    End If
  'For I% = 1 To 32767
'     ARCH$ = LUDAT$ & "TRAZA" & TRIM$(Str$(I%)) & ".TXT"
'     If EXISTE%(ARCH$) < 1 Then
'       NARCHI2% = FILEOPEN%(ARCH$, 2, 128) ' ABRE EL ARCHIVO DE ESCRITURA
'       Print #NARCHI2%, TXTRAZA$ ' ESCRIBE EN EL ARCHIVO EL NUEVO TEXTO
'       Close #NARCHI2%
'       Exit For
'     End If
'Next I%
   '
TXTRAZA$ = ""
   YAGRABOCAE% = 1
   EPAC$ = UCase$(TRIM$(EMPREAC$))

'    If InStr(EPAC$, "AHP") > 0 Then
'     If ASIGLOAHP.List2.ListCount >= 1 Then
'          ASIGLOAHP.EJECUTAR = 1
'          ASIGLOAHP.Show 1
'          ASIGLOAHP.EJECUTAR = 0
'     End If
'    End If
   
   ' ACTUALIZA EL SISTEMA BIMONETARIO QUE LLEVA EN EL ARCHIVO CARBIMON
   If ULTREG&("CARBIMON") > 0 Then
        For i& = 1 To ULTREG&("CARBIMON")
            REGISTRO$ = REGLEIDO$("CARBIMON", i&)
            Comprobante$ = Mid$(REGISTRO$, 13, 18)
            
            If Comprobante$ = NUFACTUVIE$ Then
                NumeroComprobante& = XVALO(Mid$(NUFACTUNUE$, 9, 8))
                Call REPLA(REGISTRO$, MKS(NumeroComprobante&), 7, 4)
                Call REPLA(REGISTRO$, MKS(NumeroComprobante&), 39, 4)
                
                Call REPLA(REGISTRO$, NUFACTUNUE$, 13, 18)
                Call REPLA(REGISTRO$, NUFACTUNUE$, 45, 18)
                
                Call GRAREG("CARBIMON", REGISTRO$, i&)
            End If
        Next i&
   End If
   
   Call PRINFACTU(NUFACTUNUE$, NCLI, LETRACOMP$, "PRINT")


End Sub
Function ARMACODIGOBARRA(NUFACTU$, Cae$, VENCIMIENTO$)
    
    ' Obtiene CUIT
    N1% = FILEOPEN(LUDAT$ + "SVD147.EXE", 0, 41)
    Dim x As String * 41
    Get #N1%, 6, x
    CUITT$ = TRIM$(REPLACAR$(TRIM$(DESENCRI$(x)), "-", ""))
    Close #N1%
    ' Tipo de comrpobante
    Comprobante$ = TRIM$(TIPOX_CBTEX$(NUFACTU$))
    ' Punto de venta
    PuntoVenta$ = TRIM$(Mid$(NUFACTU$, 4, 4))
    ' CAE
    CODIGOAUTORIZACION$ = TRIM$(Cae$)
    ' Fecha de vencimiento
    FechaVencimiento$ = TRIM$(VENCIMIENTO$)
    ' Arma cadena
    XX$ = CUITT$ + Comprobante$ + PuntoVenta$ + CODIGOAUTORIZACION$ + FechaVencimiento$
    ' Obtiene digito verificador
    DIGITOVERIFICADOR$ = OBTIENEDIGITO$(XX$)
    ' Arma cadena final
    XX$ = XX$ + DIGITOVERIFICADOR$
    
    YY$ = CODIBARRA$(XX$, 1) ' SE REEMPLAZA POR ESTA FUNCION
    '
    ARMACODIGOBARRA = YY$

End Function

Function ARMACODIGOQR(IdTransaccion As String)
    
    Call CIERRARCH("*.*")
    
    N1% = FILEOPEN(LUCOM$ + "facelec\cae_" & IdTransaccion & "_qr.txt", 0, 1024)
    Dim x As String * 1024
    Get #N1%, 1, x
    Close #N1%
    
    YY$ = "https://www.afip.gob.ar/fe/qr/?p=" & x
    '
    ARMACODIGOQR = YY$

End Function

Function TIPOX_CBTEX$(NUFACTU$)

    TIPOCOMP$ = Mid$(NUFACTU$, 1, 2) & ":" & Mid$(NUFACTU$, 18, 1)
    
    Select Case TIPOCOMP$
        Case "FC:A": valor$ = "01"
        Case "ND:A": valor$ = "02"
        Case "NC:A": valor$ = "03"
        Case "FC:B": valor$ = "06"
        Case "ND:B": valor$ = "07"
        Case "NC:B": valor$ = "08"
        Case "FC:E": valor$ = "19"
        Case "ND:E": valor$ = "20"
        Case "NC:E": valor$ = "21"
    End Select
    '
    TIPOX_CBTEX$ = valor$
    '
End Function
Sub PRINFACTU(NUFACTU$, NCLI, LETR$, MODO$)
    '
    'PASADA LA LETRA COMO VARIABLE POR QUE DEFINE EL FORMULARIO DE IMPRESION
    ' VER SI ESTO SE VA A CONFIGURAR
    If SOLOCAE% = 1 Then Exit Sub
    '
    ORCOP$(1) = "ORIGINAL"
    ORCOP$(2) = "DUPLICADO"
    ORCOP$(3) = "TRIPLICADO"
    ORCOP$(4) = "CUADRUPLICADO"
    ORCOP$(5) = "QUINTUPLICADO"
    ORCOP$(6) = "SEXTUPLICADO"
    '
    NC2 = NCLI
    '
    'SE AGREGA POR QUE AHORA GUARDA EL CLIENTE ORIGINAL PARA CASOS DE GRUPO DIVISIONAL
    Condi$ = "CODICOM_LAR = '" & NUFACTU$ & "'AND  NUCLIEN = " & NCLI
    NC1 = XVALO(VALOBADA("CAJABANC", "CLI_REAL", Condi$, "NOMESS"))
    If NC1 > 0 Then NC2 = NC1
    '
    If CANTCOPIAS% < 1 Then CANTCOPIAS% = XVALO(FINFACTUR07.Text2(0))
    If CANTCOPIAS% < 1 Then XVALO (CONTROL$("", "COPIAFAC"))
    If CANTCOPIAS% < 1 Then CANTCOPIAS% = 1
    '
    If CANTCOPIAS% > 6 Then
      For KKII% = 2 To 9
       ORCOP$(KKII%) = "COPIA"
      Next KKII%
    End If
    '
    TIPOCOM$ = Mid$(NUFACTU$, 1, 2)
    
    Condi$ = "DOSECLAR='" & NUFACTU$ & "' AND NUME_CLIE= " & NC2 & " AND COD_INTE > 0"
    CANTREG& = XVALO(CantRegistros("MOVISTO", Condi$, "NOMESS"))
    '
    TABLA1$ = "MOVISTO"
    CANTREGREPA& = 0
    If CANTREG& < 1 Then
        'VALIDO SI ES UNA REPARACION
        Condi1$ = "DOSECLAR='" & NUFACTU$ & "' AND NUME_CLIE= " & NC2 & ""
        CANTREGREPA& = XVALO(CantRegistros("MOVIREPA", Condi1$, "NOMESS"))
        If CANTREGREPA& > 0 Then TABLA1$ = "MOVIREPA"
    End If
    '
    'SI EL CAMPO DE Descri_Lar ES NULL O VACIO SE LE AGREGA EL DE LA DESCRIPCION
    '****************************************************************************
    SQLX$ = "UPDATE  " & TABLA1$ & " SET Descri_Lar = DESCRIP "
    SQLX$ = SQLX$ + " WHERE (Descri_Lar IS NULL OR Descri_Lar LIKE '') "
    SQLX$ = SQLX$ + " AND DOSECLAR = '" & NUFACTU$ & "' AND NUME_CLIE=" & NC2
    FLEXCONN.Execute (FILTSQL$(SQLX$))
    '
    '******************************************
    LETRA$ = Mid$(NUFACTU$, 18, 1)
    FORIPRE$ = TRIM$(CONTROL$("", "FORFAC-" + LETRA$))
    If CANTREG& = 0 And CANTREGREPA& = 0 Then TIPOCOM$ = "(T.Libre)"
    TIFADOC% = 2
    
    If Mid$(NUFACTU$, 1, 2) = "NC" Then TIFADOC% = 6
    If Mid$(NUFACTU$, 1, 2) = "ND" Then TIFADOC% = 3
    '
    If TIPOCOM$ = "(T.Libre)" Then
      If LETRA$ <> "E" Then   ' EXCLUYE FACTURAS DE EXPORTACION
        FORIPRE$ = TRIM$(CONTROL$("", "FACTEXLI"))
        If UCase$(LETRA$) = "B" Then
          FORIPRE$ = TRIM$(CONTROL$("", "FACTELIB"))
        End If
      End If
    End If
    '
    ' FACTURA EN MODO FACTURA DE CREDITO
    ModoCredito& = 0
    CondicionSql$ = "CODICOM_LAR = '" & NUFACTU$ & "' AND  NUCLIEN = " & NCLI & " AND MODOOPCIMOVIM='CR'"
    ModoCredito& = CantRegistros("CAJABANC", CondicionSql$)
    If ModoCredito& > 0 Then
        FORIPRE$ = TRIM$(CONTROL$("", "FACCRE-" + LETRA$))
    End If
    '
    'FORMULARIOS DE REPARACION
    If CANTREGREPA& > 0 And TIFADOC% = 2 Then
       FORIPRE$ = "FCR" + Mid$(FORIPRE$, 4)
    End If
    '
    '\\\OT-06-0038-WAR-01/02/06///
    FORIALT$ = ""
    If TIFADOC% = 3 Or TIFADOC% = 6 Then 'SI NC O ND
      If TIFADOC% = 3 Then ' SI NOTA DE DEBITO
        If UCase$(LETRA$) = "A" Then
          FORIALT$ = CONTROL$("", "FORNDE-A") 'ND-A
          ElseIf UCase$(LETRA$) = "B" Then
          FORIALT$ = CONTROL$("", "FORNDE-B") 'ND-B
          
          ElseIf UCase$(LETRA$) = "E" Then
          FORIALT$ = CONTROL$("", "FORNDE-E") 'ND-E
        End If
      End If
      '
      If TIFADOC% = 6 Then ' SI NOTA DE CREDITO
        If UCase$(LETRA$) = "A" Then
          FORIALT$ = CONTROL$("", "FORNCR-A") 'NC-A
          ElseIf UCase$(LETRA$) = "B" Then
          FORIALT$ = CONTROL$("", "FORNCR-B") 'NC-B
          ElseIf UCase$(LETRA$) = "E" Then
          FORIALT$ = CONTROL$("", "FORNCR-E") 'NC-E
        End If
      End If
    End If
    If TRIM$(FORIALT$) <> "" Then
      FORIPRE$ = FORIALT$
    End If
    
    ' NOTA DE DEBITO EN MODO FACTURA DE CREDITO
    If ModoCredito& > 0 And TIFADOC% = 3 Then
        FORIPRE$ = TRIM$(CONTROL$("", "NDBCRE-" + LETRA$))
    End If
    
    ' NOTA DE CREDITO EN MODO FACTURA DE CREDITO
    If ModoCredito& > 0 And TIFADOC% = 6 Then
        FORIPRE$ = TRIM$(CONTROL$("", "NCRCRE-" + LETRA$))
    End If
    
    '******************************************
    'EMP$ = "1-MDT": If CANTREG& = 0 Then EMP$ = "T-MDT"
    'ARCSQR$ = TIPOCOM$ & LETR$ & EMP$
    '
    'If EXISTE%(LUDAT$ + ARCSQR$ + ".SQR") = 0 Then
    '    Call COMUNI("Falta Archivo " & ARCSQR$ + ".SQR")
    'End If
    ARCSQR$ = FORIPRE$
    If EXISTE%(LUDAT$ + ARCSQR$ + ".SQR") = 0 Then
        Call COMUNI("Falta Archivo " & ARCSQR$ + ".SQR")
    End If
    
    '
    FILTX$ = TRIM$(NUFACTU$) & ";" & TRIM$(Str$(NC2))
    Condi$ = "DOSECLAR='" & NUFACTU$ & "' AND NUME_CLIE=" & NC2
    'SE INCORPORA PARA CONTEMPLAR A LAS REPARACIONES Y OBTENER DE MOVIREPA LA MONEDA
    TABMOV$ = "MOVISTO"
    If CANTREGREPA& > 0 Then TABMOV$ = "MOVIREPA"
    MONEMI% = XVALO(VALOBADA(TABMOV$, "MONEVAL", Condi$))
    'MONEMI% = XVALO(VALOBADA("MOVISTO", "MONEVAL", CONDI$))
    
    Condi$ = "CODICOM_LAR = '" & NUFACTU$ & "'AND  NUCLIEN = " & NCLI
    TextoQRAfip$ = TRIM$(SAFETEXT$(VALOBADA("cajabanc", "CodigoQRAfip", Condi$)))
    If TextoQRAfip$ <> "" Then
        TEXTOQR$ = TextoQRAfip$
        IMAG$ = "IMAGFACTU_QR.JPG"
        IDENTIFICADOR$ = "QR_FACTURA"
        ANCHOX% = 0
        ALTOX% = 0
        '
        On Error Resume Next
        Kill LUCOM$ & IMAG$
        On Error GoTo 0
        'RUTINA QUE LLAMA A APLICACION DE .NET PARA GENERAR LA IMAGEN DEL CODIGO QR
        QRX$ = TRIM$(IMAGENCODBARRA$(LUCOM$, IMAG$, TEXTOQR$, "QR_CODE", ANCHOX%, ALTOX%))
        Call T_Espera(2)
        'Image1.Picture = LoadPicture(LUCOM$ & IMAG$)
        Call GRABAR_QR(LUCOM$ & IMAG$, IDENTIFICADOR$)
        
        Call RecortaImagenQRFactura
    End If
    
    Timer4.Enabled = True
    ' Call STDFORM(ARCSQR$, FILTX$, MODO$)
    
End Sub
Private Sub Timer2_Timer()
   
   If CONTAD& >= INDICE& Then
       Timer2.Enabled = False
       If MODOAUTO = "NO" Then
            Call COMUNI("Proceso de Solicitud de CAE Finalizado")
       Else
            MODOAUTO = "NO" ' AVISA AL MODOAUTO="SI" QUE FINALIZO EL PROCESO
            Unload Me
       End If
   End If
   '
   If ESTADO% = 0 Then
      CONTAD& = CONTAD& + 1
      NUMEFAC$ = FACTU$(CONTAD&)
      Text5 = NUMEFAC$
      ESTADO% = 1
'      Call Command5_Click
   End If
   '
End Sub

Sub SOLICITA_CAE_BCAP_ID(IDENCAE$)
    '
 TXTRAZA$ = TXTRAZA$ & "PUNTO B.C. 57 = " & Now & vbCrLf
    Call ABRECABAN
    'CONDI$ = "IDCAE = '" & IDENCAE$ & "' And TEXTOCAE Is Not Null"
    'CONDI$ = "IDCAENUM = " & XVALO(IDENCAE$)
    Condi$ = " TEXTOCAE Is Not Null  AND IDCAENUM = " & XVALO(IDENCAE$)

    DATOX$ = VALOBADA("CAJABANC", "TEXTOCAE", Condi$, "NOMESS")
    If DATOX$ = "" Then
       Call COMUNI("No Se Han Grabado Correctamente Los Datos Para Solicitar el CAE a la A.F.I.P.\Capture esta Pantalla y Consulte a su Soporte de Sistemas")
       Exit Sub
    End If
TXTRAZA$ = TXTRAZA$ & "PUNTO 58= " & Now & vbCrLf    '
    'FECHA DE HOY
    HOII% = HOY(HOS$)
    FECH1$ = ARMAFECHCAE(HOII%) ' FORMATO AAAAMMDD

    FE% = CVINT(VALOBADA("CAJABANC", "FECHEMISI", Condi$, "NOMESS"))
    FECH2$ = ARMAFECHCAE(FE%) ' FORMATO AAAAMMDD
    If FECH1$ <> FECH2$ Then DATOX$ = REPLACAR(DATOX$, FECH2$, FECH1$) 'REEMPLAZO LA FECHA QUE TENIA POR LA DE HOY
TXTRAZA$ = TXTRAZA$ & "PUNTO 59 = " & Now & vbCrLf
    '
    CARPETOK$ = "C:\FLEXOFT\FACELEC\"
    Archiv$ = CARPETOK$ + "CONSULTA_CAE_" + IDENCAE$ + ".TXT"
    NARCHI2% = FILEOPEN%(Archiv$, 2, 128) ' ABRE EL ARCHIVO DE ESCRITURA
    Print #NARCHI2%, DATOX$ ' ESCRIBE EN EL ARCHIVO EL NUEVO TEXTO
    Close #NARCHI2%
TXTRAZA$ = TXTRAZA$ & "PUNTO 60 = " & Now & vbCrLf
'
    On Error Resume Next
    Kill CARPETOK$ + "CAE_" + IDENCAE$ + ".TXT"  ' ELIMINAMOS EL ARCHIVO
    On Error GoTo 0
    '
TXTRAZA$ = TXTRAZA$ & "PUNTO 61 = " & Now & vbCrLf
    'CONDI$ = "IDCAE = '" & IDENCAE$ & "'"
    Condi$ = "IDCAENUM = " & XVALO(IDENCAE$)
    PVTA$ = VALOBADA("CAJABANC", "NUMETALO", Condi$, "NOMESS")
    Condi1$ = Condi$ + "AND TIPOMOVIM = 'FVEN'"
    PREFI$ = VALOBADA("CAJABANC", "OPCIMOVIM", Condi1$, "NOMESS")
    LETRA$ = VALOBADA("CAJABANC", "VARIMOVIM", Condi1$, "NOMESS")
    '
    If EXISTE%(CARPETOK$ + "CAE_" + IDENCAE$ + ".TXT") = 0 Then
      If EXISTE%(CARPETOK$ + "wsbfev1_cliente.exe") = 1 Then
         'TC = TIPO_CBTE(PREFI$, LETRA$)
         'XXX = Shell(CARPETOK$ + "wsbfev1_cliente.exe AUTO " & IDENCAE$ & " " & TC & " " & PVTA$, 4)
         XXX = Shell(CARPETOK$ + "wsbfev1_cliente.exe AUTO " & IDENCAE$, 4)
         Flag% = 1
      ElseIf EXISTE%(CARPETOK$ + "wsbfe_cliente.exe") = 1 Then
         TC = TIPO_CBTE(PREFI$, LETRA$)
         XXX = Shell(CARPETOK$ + "wsbfe_cliente.exe AUTO " & IDENCAE$ & " " & TC & " " & PVTA$, 4)
         Flag% = 1
      Else
         Call COMUNI("Falta Ejecutable: " & REPLACAR$(CARPETOK$, "\", "/") & "wsbfev1_cliente.exe")
         Exit Sub
      End If
    End If
    '
TXTRAZA$ = TXTRAZA$ & "PUNTO 62 = " & Now & vbCrLf
End Sub

Sub SOLICITA_CAE_EXPO_ID(IDENCAE$)
TXTRAZA$ = TXTRAZA$ & "PUNTO EXPO 57 = " & Now & vbCrLf
'
    HOII% = HOY(HOS$)
    FECH1$ = ARMAFECHCAE(HOII%)
    
    Call ABRECABAN
    'CONDI$ = "IDCAE = '" & IDENCAE$ & "' And TEXTOCAE Is Not Null"
    Condi$ = " TEXTOCAE Is Not Null  AND IDCAENUM = " & XVALO(IDENCAE$)
    DATOX$ = VALOBADA("CAJABANC", "TEXTOCAE", Condi$, "NOMESS")
    FE% = CVINT(VALOBADA("CAJABANC", "FECHEMISI", Condi$, "NOMESS"))
    FECH2$ = ARMAFECHCAE(FE%)
    If FECH1$ <> FECH2$ Then DATOX$ = REPLACAR(DATOX$, FECH2$, FECH1$)
TXTRAZA$ = TXTRAZA$ & "PUNTO 58 = " & Now & vbCrLf    '
    If DATOX$ = "" Then
       Call COMUNI("No Se Han Grabado Correctamente Los Datos Para Solicitar el CAE a la A.F.I.P.\Capture esta Pantalla y Consulte a su Soporte de Sistemas")
       Exit Sub
    End If    '
    CARPETOK$ = "C:\FLEXOFT\FACELEC\"
    Archiv$ = CARPETOK$ + "CONSULTA_CAE_" + IDENCAE$ + ".TXT"
    NARCHI2% = FILEOPEN%(Archiv$, 2, 128) ' ABRE EL ARCHIVO DE ESCRITURA
    Print #NARCHI2%, DATOX$ ' ESCRIBE EN EL ARCHIVO EL NUEVO TEXTO
    Close #NARCHI2%
TXTRAZA$ = TXTRAZA$ & "PUNTO 59 = " & Now & vbCrLf    '
    On Error Resume Next
    Kill CARPETOK$ + "CAE_" + IDENCAE$ + ".TXT"  ' ELIMINAMOS EL ARCHIVO
    On Error GoTo 0
TXTRAZA$ = TXTRAZA$ & "PUNTO 60 = " & Now & vbCrLf
'
    If EXISTE%(CARPETOK$ + "CAE_" + IDENCAE$ + ".TXT") = 0 Then
      
      If EXISTE%(CARPETOK$ + "wsfexv1_cliente.exe") = 1 Then
         XXX = Shell(CARPETOK$ + "wsfexv1_cliente.exe AUTO " & IDENCAE$, 4)
         Flag% = 1
         Exit Sub
      End If
    
      If EXISTE%(CARPETOK$ + "wsfex_cliente.exe") = 1 Then
         XXX = Shell(CARPETOK$ + "wsfex_cliente.exe AUTO " & IDENCAE$, 4)
         Flag% = 1
      Else
         Call COMUNI("Falta Ejecutable: " & REPLACAR$(CARPETOK$, "\", "/") & "wsfex_cliente.exe")
         Exit Sub
      End If
    End If
TXTRAZA$ = TXTRAZA$ & "PUNTO 61 = " & Now & vbCrLf    '
End Sub

Sub SOLICITA_CAE_ID(IDENCAE$, Optional ModoCredito%)
    '
 TXTRAZA$ = TXTRAZA$ & "PUNTO CAE 57 = " & Now & vbCrLf
    
    Call ABRECABAN
    'CONDI$ = "IDCAE = '" & IDENCAE$ & "' And TEXTOCAE Is Not Null"
    'CONDI$ = "IDCAENUM = " & XVALO(IDENCAE$)
    Condi$ = " TEXTOCAE Is Not Null  AND IDCAENUM = " & XVALO(IDENCAE$)

    DATOX$ = VALOBADA("CAJABANC", "TEXTOCAE", Condi$, "NOMESS")
    If DATOX$ = "" Then
       Call COMUNI("No Se Han Grabado Correctamente Los Datos Para Solicitar el CAE a la A.F.I.P.\Capture esta Pantalla y Consulte a su Soporte de Sistemas")
       Exit Sub
    End If
    '
TXTRAZA$ = TXTRAZA$ & "PUNTO 58 = " & Now & vbCrLf
    'FECHA DE HOY
    HOII% = HOY(HOS$)
    FECH1$ = ARMAFECHCAE(HOII%) ' FORMATO AAAAMMDD

    FE% = CVINT(VALOBADA("CAJABANC", "FECHEMISI", Condi$, "NOMESS"))
    FECH2$ = ARMAFECHCAE(FE%) ' FORMATO AAAAMMDD
    If FECH1$ <> FECH2$ Then DATOX$ = REPLACAR(DATOX$, FECH2$, FECH1$)  'REEMPLAZO LA FECHA QUE TENIA POR LA DE HOY
'    MsgBox "paso"
'    Comprobante$ = VALOBADA("CAJABANC", "CODICOM_LAR", Condi$, "NOMESS")
'    ultnum_tipo$ = Mid$(Comprobante$, 1, 2): ultnum_letra$ = Mid$(Comprobante$, 18, 1): ultnum_talonario% = CInt(Mid$(Comprobante$, 4, 4))
'    UltimoNumeroComprobante& = XVALO(VALOBADA("cajabanc", "max(numecomp)", "tipomovim='" & ultnum_tipo$ & " and varimovim='" & ultnum_letra$ & "' and numetalo=" & CStr(ultnum_talonario%) & " and tipomovim='fven'"))
'    If UltimoNumeroComprobante& > 0 Then
'        DATOX$ = REPLACAR(DATOX$, "<cbt_desde>0</cbt_desde>", "<cbt_desde>" & CStr(UltimoNumeroComprobante&) & "</cbt_desde>")
'        MsgBox DATOX$
'    End If

    CARPETOK$ = "C:\FLEXOFT\FACELEC\"
    Archiv$ = CARPETOK$ + "CONSULTA_CAE_" + IDENCAE$ + ".TXT"
    NARCHI2% = FILEOPEN%(Archiv$, 2, 128) ' ABRE EL ARCHIVO DE ESCRITURA
    Print #NARCHI2%, DATOX$ ' ESCRIBE EN EL ARCHIVO EL NUEVO TEXTO
    Close #NARCHI2%
    '
 TXTRAZA$ = TXTRAZA$ & "PUNTO 59 = " & Now & vbCrLf
    On Error Resume Next
    Kill CARPETOK$ + "CAE_" + IDENCAE$ + ".TXT"  ' ELIMINAMOS EL ARCHIVO
    On Error GoTo 0
    '
 TXTRAZA$ = TXTRAZA$ & "PUNTO 60 = " & Now & vbCrLf
 
    If EXISTE%(CARPETOK$ + "CAE_" + IDENCAE$ + ".TXT") = 0 Then
           
      If EXISTE%(CARPETOK$ + "wsfev1_cliente.exe") = 1 Then
         XXX = Shell(CARPETOK$ + "wsfev1_cliente.exe AUTO " & IDENCAE$, 4)
         Flag% = 1
         Exit Sub
      End If
    
      If EXISTE%(CARPETOK$ + "wsfev2_cliente.exe") = 1 And ModoCredito = 1 Then
         XXX = Shell(CARPETOK$ + "wsfev2_cliente.exe AUTO " & IDENCAE$, 4)
         Flag% = 1
         Exit Sub
      End If
    
      If EXISTE%(CARPETOK$ + "wsfe_cliente.exe") = 1 Then
         XXX = Shell(CARPETOK$ + "wsfe_cliente.exe AUTO " & IDENCAE$, 4)
         Flag% = 1
      Else
         Call COMUNI("Falta Ejecutable: " & REPLACAR$(CARPETOK$, "\", "/") & "wsfe_cliente.exe")
         Exit Sub
      End If
    End If
    '
TXTRAZA$ = TXTRAZA$ & "PUNTO 61= " & Now & vbCrLf
End Sub
Function ARMAFECHCAE(FE%)

    If FE% < 1 Then Exit Function
    
    fech = FECHATEXLAR$(FE%)
    FECHAFORM = Right(fech, 4) + Mid$(fech, 4, 2) + Mid$(fech, 1, 2)
    ARMAFECHCAE = FECHAFORM
    '
End Function

Private Sub Timer3_Timer()
'''
''    Timer3.Enabled = False
''    Call Command2_Click
'''
End Sub

Private Sub Timer4_Timer()
    'Static TOTACOPIAS%
    '
    If PORTUFAC$ <> "" Then
        Call SETPRINPORT(PORTUFAC$)
    End If
    
    'LEE LA CANTIDAD DE COPIAS
10  If TOTACOPIAS% < 1 Then TOTACOPIAS% = XVALO(FINFACTUR07.Text2(0))
     
    If TOTACOPIAS% < 1 Then TOTACOPIAS% = XVALO(CONTROL$("", "COPIAFAC"))
    CANTCOPIAS% = CANTCOPIAS% - 1
    'AGREGA EL TEXTO SEGUN COPIA (ORIGINAL, DUPLICADO,ETC)
    '
    'IMPRIME
    ENVIAMAILFACTURA% = FINFACTUR07.Check1.Value 'SI ESTA TILDADO ENVIA POR MAIL
    Call STDFORM(ARCSQR$, FILTX$ & ";" & ORCOP$(TOTACOPIAS% - (CANTCOPIAS%)) & ";" & "QR_FACTURA", "PRINT", , ENVIAMAILFACTURA%)
    FINFACTUR07.Check1.Value = 0 ' CANCELA PARA QUE NO VUELVA A ENVIAR EL MAIL
    '
    'MANTENERLO EN OBSERVACION !!!!!!!
    If CANTCOPIAS% < 1 Then
         'SE INCLUYE ESTE CODIGO PARA PERMITIR REPETIR LA IMPRESION LUEGO DE
         'IMPRIMIRSE LAS COPIAS POR DEFAULT ESTABLECIDAS
         TTXYZ$ = "Se ha Emitido por Impresora"
         NN = InStr(1, FILTX$, ";"): NUMFAC$ = Mid$(FILTX$, 1, NN - 1)
         TTXYZ$ = TTXYZ$ + "\Comprobante Numero '" + NUMFAC$ + "'"
         If COMALTER%(TTXYZ$ + "\\Repetir Impresion\Continuar") = 1 Then
            CANTCOPIAS% = TOTACOPIAS%: GoTo 10 'SI REPITE VUELVA A IMPRIMIR LA CANTIDAD COFIGURADA
         End If
         Timer4.Enabled = False
    End If
End Sub

Private Sub tmrRecuperaComprobanteDesdeAfip_Timer()
    Static CONTA%
    Static EnProceso As Boolean
    Static CantidadDeProcesados As Integer
    Static Inicializa As Byte
    Static ComprobanteEnRecuperacion As String
    Static PuntoVentaComun As Long
    Static PuntoVentaCredito As Long
    Static PuntoVentaBienes As Long
    Static PuntoVentaExportacion As Long
    Static FechaCbteRecuperacionManual As String
    Static CuitCbteRecuperacionManual As String
    Static ImporteTotalCbteRecuperacionManual As String
    
    tmrRecuperaComprobanteDesdeAfip.Enabled = False
    Screen.MousePointer = 11
    
    CARPETOK$ = "C:\FLEXOFT\FACELEC\"
    TX1$ = ""
    IDENCAE$ = TRIM$(Text5)
    Cae$ = ""
    ImporteTotal$ = ""
    NumeroComprobante$ = ""
    CuitCliente$ = ""
    PuntoVenta$ = ""
    VencimientoCae$ = ""
    TipoComprobante$ = ""
    FechaEmision$ = ""
    '
    If Inicializa = 0 Then
        EnProceso = False
        CantidadDeProcesados = 1
        Campos$ = "Fecha/D10;Comprobante/A18;Cuit/A13;Cuenta/I6;Cliente/A36;Importe/G14.2;Cae/A20;Vto/A10;"
        Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
        Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
        FRMZELECHO.MSF2.Rows = 1
        
        PuntoVentaComun = XVALO(CONTROL("FACTURA", "PUVENELE"))
        PuntoVentaBienes = XVALO(CONTROL("FACTURA", "PUVENCAP"))
        PuntoVentaExportacion = XVALO(CONTROL("FACTURA", "PUVEFAEX"))
        PuntoVentaCredito = XVALO(CONTROL("FACTURA", "PUVECRED"))
        
        Text3 = Text3 & vbCrLf & "Recuperando comprobante desde Afip ..."
                
        Inicializa = 1
    End If
    '
    If EnProceso = False Then   ' PROXIMO HUECO A RECUPERAR SUS DATOS
        
        If CantidadDeProcesados <= UBound(RecuperacionHuecos$) Then
            
            ComprobanteEnRecuperacion = RecuperacionHuecos$(CantidadDeProcesados)
            If ModoSolicitudManual = True Then
                CbteRecuperacionManual$ = TRIM$(Text6)
                    
                    CondicionSql$ = "numecomp<0 and tipomovim='fven' " _
                    & " and cli_real>0" _
                    & " and opcimovim='" & Mid$(CbteRecuperacionManual$, 1, 2) & "'" _
                    & " and varimovim='" & Mid$(CbteRecuperacionManual$, 18, 1) & "'" _
                    & " and numetalo=" & XVALO(Mid$(REPLACAR$(CbteRecuperacionManual$, "X", "0"), 4, 4))
                
                    fechaCbte% = CVINT(VALOBADA("cajabanc", "fechconta", CondicionSql$))
                FechaCbteRecuperacionManual$ = FECHATEX$(fechaCbte%)
                    clienteCbte& = VALOBADA("cajabanc", "nuclien", CondicionSql$)
                CuitCbteRecuperacionManual$ = REPLACAR$(CuitCli$(clienteCbte&), "-", "")
                    importeCbte# = XVALO(VALOBADA("cajabanc", "valabscomp", CondicionSql$))
                ImporteTotalCbteRecuperacionManual$ = TRIM$(FORMATNUM$(importeCbte#, "F14.2"))
            End If
            Text3 = Text3 & vbCrLf & "Recuperando '" & ComprobanteEnRecuperacion & "' ..."
            
            Select Case RecuperacionPuntoVenta%
                Case PuntoVentaComun:   DoEvents: Call T_Espera(1)
                                        Ejecuta$ = "c:\flexoft\facelec\wsfev1_cliente.exe RECUP " & ComprobanteEnRecuperacion & " COMUN"
                                        If RecuperacionModoCredito = True Then Ejecuta$ = "c:\flexoft\facelec\wsfev1_cliente.exe RECUP " & ComprobanteEnRecuperacion & " CREDITO"
                                        r = Shell(Ejecuta$, 4)
                
                Case PuntoVentaCredito: DoEvents: Call T_Espera(1)
                                        r = Shell("c:\flexoft\facelec\wsfev1_cliente.exe RECUP " & ComprobanteEnRecuperacion & " CREDITO", 4)
                
                Case PuntoVentaBienes:  DoEvents: Call T_Espera(1)
                                        r = Shell("c:\flexoft\facelec\wsbfev1_cliente.exe RECUP " & ComprobanteEnRecuperacion, 4)
                
                Case PuntoVentaExportacion: DoEvents: Call T_Espera(1)
                                            r = Shell("c:\flexoft\facelec\wsfexv1_cliente.exe RECUP " & ComprobanteEnRecuperacion, 4)
            End Select

            EnProceso = True
            tmrRecuperaComprobanteDesdeAfip.Enabled = True
            
        Else
        
            ' PRESENTA LISTA DE HUECOS A RECUPERAR
            Screen.MousePointer = 1
            FRMZELECHO.Caption = "Comprobantes Registrados En AFIP Pendientes de Recuperar"
            FRMZELECHO.Width = 12500
10          FRMZELECHO.Show 1
            
            Cae$ = TRIM$(RESP_ZELE_VECT(7))
            If Cae$ <> "" Then
                
                ComprobanteARecuperar$ = TRIM$(RESP_ZELE_VECT(2))
                FechaEmision$ = TRIM$(RESP_ZELE_VECT(1))
                NroCliente& = 0
                CuitCliente$ = TRIM$(RESP_ZELE_VECT(3))
                ImporteComprobante# = XVALO(RESP_ZELE_VECT(6))
                VencimientoCae$ = TRIM$(RESP_ZELE_VECT(8))
                
                If BuscaComprobanteAfipEnSistema(ComprobanteARecuperar$, FechaEmision$, CuitCliente$, ImporteComprobante#, Cae$, VencimientoCae$, NroCliente&) = True Then
                    If FRMZELECHO.MSF2.Rows > 2 Then
                        For i = 1 To FRMZELECHO.MSF2.Rows - 1
                            If FRMZELECHO.MSF2.TextMatrix(i, 7) = Cae$ Then
                                FRMZELECHO.MSF2.RemoveItem (i)
                                FRMZELECHO.MSF2.Refresh
                                Exit For
                            End If
                        Next i
                    Else
                        FRMZELECHO.MSF2.Rows = 1
                    End If
                
                Else
                    Screen.MousePointer = 1
                    If NroCliente& > 0 Then
                        RecuperaTextoLibre% = COMALTER%("No se Pudo Recuperar Comprobante '" & ComprobanteARecuperar$ & "'.\¿Desea Agregarlo Al Sistema?\\No, Continuar\Si, Agregar")
    
                        If RecuperaTextoLibre% = 2 Then
                            Screen.MousePointer = 11
                            Call GrabaComprobanteRecuperado(ComprobanteARecuperar$, NroCliente&)
                            
                            For i = 1 To FRMZELECHO.MSF2.Rows - 1
                                If FRMZELECHO.MSF2.TextMatrix(i, 7) = Cae$ Then
                                    If FRMZELECHO.MSF2.Rows - 1 > 1 Then
                                        FRMZELECHO.MSF2.RemoveItem (i)
                                    Else
                                        FRMZELECHO.MSF2.Rows = 1
                                    End If
                                    FRMZELECHO.MSF2.Refresh
                                    Exit For
                                End If
                            Next i
                            Screen.MousePointer = 1
                        End If
                    End If
                End If
                
                If ModoSolicitudManual = False Then GoTo 10
            Else
                Text3 = ""
                If ModoSolicitudManual = True Then
                    OpcionSolicitarCae% = COMALTER%("¿Desea Solicitar Nuevamente Autorización a AFIP?\\No, Continuar\Si, Solicitar")
                    If OpcionSolicitarCae% = 2 Then
                        If CantRegistros("cajabanc", "codicom_lar='" & Text6 & "' and tipomovim='fven'") > 0 Then
                            Call SOLICITACAE
                        End If
                    End If
                    ModoSolicitudManual = False
                    Command4.Enabled = True
                End If
            End If
            
            Inicializa = 0
            
        End If
        
    Else    ' EN PROCESO DE RECUPERACION DE DATOS
    
        ' LEE ARCHIVO CON LOS DATOS RECUPERADOS DESDE AFIP
        If EXISTE%(CARPETOK$ + "CAE_" + ComprobanteEnRecuperacion + ".TXT") > 0 Then
            NARCHIVO% = FILEOPEN%(CARPETOK$ + "CAE_" + ComprobanteEnRecuperacion + ".TXT", 1, 128)  'ABRE EL ARCHIVO DE LECTURA
            Do While Not EOF(NARCHIVO%) ' RECORRE
              Line Input #NARCHIVO%, XXXXX$ ' LEE CADA REGISTRO
              TX1$ = TX1$ + XXXXX$ + vbCrLf
              If InStr(UCase$(XXXXX$), UCase$("<CodAutorizacion>")) > 0 Then    ' CAE
                 POS1% = InStr(UCase$(XXXXX$), UCase$("<CodAutorizacion>")) + 17
                 TXT$ = Mid$(XXXXX$, POS1%)
                 POS2% = InStr(TXT$, "</")
                 Cae$ = Mid$(TXT$, 1, POS2% - 1)
              ElseIf InStr(UCase$(XXXXX$), UCase$("<Cae>")) > 0 Then    ' CAE
                 POS1% = InStr(UCase$(XXXXX$), UCase$("<Cae>")) + 5
                 TXT$ = Mid$(XXXXX$, POS1%)
                 POS2% = InStr(TXT$, "</")
                 Cae$ = Mid$(TXT$, 1, POS2% - 1)
              ElseIf InStr(UCase$(XXXXX$), UCase$("<ImpTotal>")) > 0 Then   ' IMPORTE TOTAL
                 POS1% = InStr(UCase$(XXXXX$), UCase$("<ImpTotal>")) + 10
                 TXT$ = Mid$(XXXXX$, POS1%)
                 POS2% = InStr(TXT$, "</")
                 ImporteTotal$ = Mid$(TXT$, 1, POS2% - 1)
              ElseIf InStr(UCase$(XXXXX$), UCase$("<Imp_total>")) > 0 Then   ' IMPORTE TOTAL
                 POS1% = InStr(UCase$(XXXXX$), UCase$("<Imp_total>")) + 11
                 TXT$ = Mid$(XXXXX$, POS1%)
                 POS2% = InStr(TXT$, "</")
                 ' SE CONDICIONA A QUE ESTE VACIO, PORQUE LA ETIQUETA Imp_total ESTA DOS VECES EN EL ARCHIVO DEVUELTA.
                 ' LA PRIMERA APARICION ES LA QUE TIENE EL VALOR TOTAL DEL COMPROBANTE.
                 If ImporteTotal$ = "" Then ImporteTotal$ = Mid$(TXT$, 1, POS2% - 1)
              ElseIf InStr(UCase$(XXXXX$), UCase$("<CbteDesde>")) > 0 Then  ' NUMERO COMPROBANTE
                 POS1% = InStr(UCase$(XXXXX$), UCase$("<CbteDesde>")) + 11
                 TXT$ = Mid$(XXXXX$, POS1%)
                 POS2% = InStr(TXT$, "</")
                 NumeroComprobante$ = Mid$(TXT$, 1, POS2% - 1)
              ElseIf InStr(UCase$(XXXXX$), UCase$("<Cbte_nro>")) > 0 Then  ' NUMERO COMPROBANTE
                 POS1% = InStr(UCase$(XXXXX$), UCase$("<Cbte_nro>")) + 10
                 TXT$ = Mid$(XXXXX$, POS1%)
                 POS2% = InStr(TXT$, "</")
                 NumeroComprobante$ = Mid$(TXT$, 1, POS2% - 1)
              ElseIf InStr(UCase$(XXXXX$), UCase$("<DocNro>")) > 0 Then     ' CUIT
                 POS1% = InStr(UCase$(XXXXX$), UCase$("<DocNro>")) + 8
                 TXT$ = Mid$(XXXXX$, POS1%)
                 POS2% = InStr(TXT$, "</")
                 CuitCliente$ = Mid$(TXT$, 1, POS2% - 1)
              ElseIf InStr(UCase$(XXXXX$), UCase$("<Nro_doc>")) > 0 Then     ' CUIT
                 POS1% = InStr(UCase$(XXXXX$), UCase$("<Nro_doc>")) + 9
                 TXT$ = Mid$(XXXXX$, POS1%)
                 POS2% = InStr(TXT$, "</")
                 CuitCliente$ = Mid$(TXT$, 1, POS2% - 1)
              ElseIf InStr(UCase$(XXXXX$), UCase$("<Cuit_pais_cliente>")) > 0 Then     ' CUIT
                 POS1% = InStr(UCase$(XXXXX$), UCase$("<Cuit_pais_cliente>")) + 19
                 TXT$ = Mid$(XXXXX$, POS1%)
                 POS2% = InStr(TXT$, "</")
                 CuitCliente$ = Mid$(TXT$, 1, POS2% - 1)
              ElseIf InStr(UCase$(XXXXX$), UCase$("<PtoVta>")) > 0 Then     ' PUNTO DE VENTA
                 POS1% = InStr(UCase$(XXXXX$), UCase$("<PtoVta>")) + 8
                 TXT$ = Mid$(XXXXX$, POS1%)
                 POS2% = InStr(TXT$, "</")
                 PuntoVenta$ = Mid$(TXT$, 1, POS2% - 1)
              ElseIf InStr(UCase$(XXXXX$), UCase$("<Punto_vta>")) > 0 Then     ' PUNTO DE VENTA
                 POS1% = InStr(UCase$(XXXXX$), UCase$("<Punto_vta>")) + 11
                 TXT$ = Mid$(XXXXX$, POS1%)
                 POS2% = InStr(TXT$, "</")
                 PuntoVenta$ = Mid$(TXT$, 1, POS2% - 1)
              ElseIf InStr(UCase$(XXXXX$), UCase$("<FchVto>")) > 0 Then     ' VENCIMIENTO CAE
                 POS1% = InStr(UCase$(XXXXX$), UCase$("<FchVto>")) + 8
                 TXT$ = Mid$(XXXXX$, POS1%)
                 POS2% = InStr(TXT$, "</")
                 VencimientoCae$ = Mid$(TXT$, 1, POS2% - 1)
              ElseIf InStr(UCase$(XXXXX$), UCase$("<Fch_venc_Cae>")) > 0 Then     ' VENCIMIENTO CAE
                 POS1% = InStr(UCase$(XXXXX$), UCase$("<Fch_venc_Cae>")) + 14
                 TXT$ = Mid$(XXXXX$, POS1%)
                 POS2% = InStr(TXT$, "</")
                 VencimientoCae$ = Mid$(TXT$, 1, POS2% - 1)
              ElseIf InStr(UCase$(XXXXX$), UCase$("<CbteTipo>")) > 0 Then   ' TIPO COMPROBANTE
                 POS1% = InStr(UCase$(XXXXX$), UCase$("<CbteTipo>")) + 10
                 TXT$ = Mid$(XXXXX$, POS1%)
                 POS2% = InStr(TXT$, "</")
                 TipoComprobante$ = Mid$(TXT$, 1, POS2% - 1)
              ElseIf InStr(UCase$(XXXXX$), UCase$("<Tipo_cbte>")) > 0 Then   ' TIPO COMPROBANTE
                 POS1% = InStr(UCase$(XXXXX$), UCase$("<CbteTipo>")) + 11
                 TXT$ = Mid$(XXXXX$, POS1%)
                 POS2% = InStr(TXT$, "</")
                 TipoComprobante$ = Mid$(TXT$, 1, POS2% - 1)
              ElseIf InStr(UCase$(XXXXX$), UCase$("<CbteFch>")) > 0 Then    ' FECHA
                 POS1% = InStr(UCase$(XXXXX$), UCase$("<CbteFch>")) + 9
                 TXT$ = Mid$(XXXXX$, POS1%)
                 POS2% = InStr(TXT$, "</")
                 FechaEmision$ = Mid$(TXT$, 1, POS2% - 1)
              ElseIf InStr(UCase$(XXXXX$), UCase$("<Fecha_cbte_orig>")) > 0 Then    ' FECHA
                 POS1% = InStr(UCase$(XXXXX$), UCase$("<Fecha_cbte_orig>")) + 17
                 TXT$ = Mid$(XXXXX$, POS1%)
                 POS2% = InStr(TXT$, "</")
                 FechaEmision$ = Mid$(TXT$, 1, POS2% - 1)
              ElseIf InStr(UCase$(XXXXX$), UCase$("<Fecha_cbte>")) > 0 Then    ' FECHA
                 POS1% = InStr(UCase$(XXXXX$), UCase$("<Fecha_cbte>")) + 12
                 TXT$ = Mid$(XXXXX$, POS1%)
                 POS2% = InStr(TXT$, "</")
                 FechaEmision$ = Mid$(TXT$, 1, POS2% - 1)
              ElseIf InStr(UCase$(XXXXX$), UCase$("<Obs>")) > 0 Then    ' OBS (CLIENTE PARA EXPORTACION)
                 POS1% = InStr(UCase$(XXXXX$), UCase$("<Obs>")) + 5
                 TXT$ = Mid$(XXXXX$, POS1%)
                 POS2% = InStr(TXT$, "</")
                 If POS2% > POS1% Then
                    N0% = InStr(1, Mid$(TXT$, 1, POS2% - 1), "(")
                    N1% = InStr(1, Mid$(TXT$, 1, POS2% - 1), ")")
                    If N0% > 0 And N1% And N1% > N0% Then
                       NumeroCliente& = XVALO(Mid$(Mid$(TXT$, 1, POS2% - 1), N0% + 1, N1% - N0% - 1))
                    End If
                 End If
              End If
            Loop
            Close NARCHIVO%
        End If
        '
        If XVALO(Cae$) > 0 Then
        
            ' COMPLETA LISTA CON LOS DATOS DEL HUECO (COMPROBANTE) A RECUPERAR
            FRMZELECHO.MSF2.Rows = FRMZELECHO.MSF2.Rows + 1
            J& = FRMZELECHO.MSF2.Rows - 1
                FechaFormateada$ = Mid$(FechaEmision$, 7, 2) & "/" & Mid$(FechaEmision$, 5, 2) & "/" & Mid$(FechaEmision$, 3, 2)
            FRMZELECHO.MSF2.TextMatrix(J&, 1) = FechaFormateada$
            FRMZELECHO.MSF2.TextMatrix(J&, 2) = ComprobanteEnRecuperacion
            FRMZELECHO.MSF2.TextMatrix(J&, 3) = CuitCliente$
                Cuit$ = Mid$(CuitCliente$, 1, 2) & "-" & Mid$(CuitCliente$, 3, 8) & "-" & Mid$(CuitCliente$, 11, 1)
                CondicionSql$ = "(cuit_cli='" & Cuit$ & "' or cuit_cli='" & CuitCliente$ & "') and stat_cli>0 and gru_div=0"
                If XVALO(Mid$(CuitCliente$, 1, 2)) >= 50 Then
                    CondicionSql$ = "nume_cli=" & CStr(NumeroCliente&)
                End If
                Cliente$ = VALOBADA("deudor12", "raso_cli", CondicionSql$)
                NumeroCliente& = XVALO(VALOBADA("deudor12", "nume_cli", CondicionSql$))
            FRMZELECHO.MSF2.TextMatrix(J&, 4) = CStr(NumeroCliente&)
            FRMZELECHO.MSF2.TextMatrix(J&, 5) = Cliente$
            FRMZELECHO.MSF2.TextMatrix(J&, 6) = FORMATNUM$(XVALO(ImporteTotal$), "F14.2")
            FRMZELECHO.MSF2.TextMatrix(J&, 7) = Cae$
            FRMZELECHO.MSF2.TextMatrix(J&, 8) = VencimientoCae$
            
            If ModoSolicitudManual = True Then
                If FechaCbteRecuperacionManual$ = FechaFormateada$ And _
                    CuitCbteRecuperacionManual$ = CuitCliente$ And _
                    XVALO(ImporteTotalCbteRecuperacionManual$) = XVALO(ImporteTotal$) Then
                    
                    Call COLOREAR_GRILLA_SOLAFILA(FRMZELECHO.MSF2, RGB(201, 184, 143), J&)
                End If
            End If
            
            CONTA% = 0
            CantidadDeProcesados = CantidadDeProcesados + 1
            EnProceso = False
            tmrRecuperaComprobanteDesdeAfip.Enabled = True
            
        Else
        
            CONTA% = CONTA% + 1
            If CONTA% < 40 Then ' CANTIDAD DE INTENTOS HASTA ESPERAR LOS DATOS DEL HUECO(COMPROBANTE) EN RECUPERACION
                tmrRecuperaComprobanteDesdeAfip.Enabled = True
            Else
                CONTA% = 0
                CantidadDeProcesados = CantidadDeProcesados + 1
                EnProceso = False
                tmrRecuperaComprobanteDesdeAfip.Enabled = True
            End If
            
        End If
        
    End If
    '
End Sub

Sub GrabaComprobanteRecuperado(Comprobante$, NumeroCliente&)
 
    CARPETOK$ = "C:\FLEXOFT\FACELEC\"
    '
    Dim CUEIVX$(2)
    CONT% = 0
    TotalComprobante# = 0
    ImporteNeto# = 0
    ImporteIva# = 0
    ImporteExento# = 0
    ImporteNoGravado# = 0
    Moneda$ = ""
    TipoCambio# = 0
    IDECAE$ = "0"
    PuntoVenta% = XVALO(Mid$(Comprobante$, 4, 4))
    '
    ImportePercepcionIB# = 0
    BasePercepcionIBCaba# = 0
    BasePercepcionIBBsAs# = 0
    BasePercepcionIBMisiones# = 0
    PercepcionIBCaba# = 0
    PercepcionIBBsAs# = 0
    PercepcionIBMisiones# = 0
    '
    ImporteIva# = 0
    IvaNoGravado# = 0
    IvaExento# = 0
    IvaTasa0# = 0
    BaseImponibleNoGravada# = 0
    BaseImponibleExenta# = 0
    IvaTasa105# = 0: BaseImponible105# = 0
    IvaTasa210# = 0: BaseImponible210# = 0
    IvaTasa270# = 0: BaseImponible270# = 0
    '
    NARCHIVO% = FILEOPEN%(CARPETOK$ + "CAE_" + Comprobante$ + ".TXT", 1, 128) 'ABRE EL ARCHIVO DE LECTURA
    '
    '********************
    'PUNTO DE VENTA COMUN
    '********************
    If PuntoVenta% = XVALO(CONTROL("FACTURA", "PUVENELE")) Or PuntoVenta% = XVALO(CONTROL("FACTURA", "PUVECRED")) Then '
         Dim doc As New MSXML2.DOMDocument
         Dim success As Boolean

         success = doc.Load(CARPETOK$ + "CAE_" + Comprobante$ + ".TXT")

         Set nodeListEncabezado = doc.selectNodes("/FECompConsultaResponse/ResultGet")
         If Not nodeListEncabezado Is Nothing Then
            Dim node As MSXML2.IXMLDOMNode
            Dim nodeh As MSXML2.IXMLDOMNode

            For Each node In nodeListEncabezado
                'NumeroCliente& = OPVARFA07.CLIENTCUIT&(node.selectSingleNode("DocNro").TEXT)
                FechaEmision% = OPVARFA07.INVERFECHA%(node.selectSingleNode("CbteFch").TEXT)
                TotalComprobante# = XVALO(node.selectSingleNode("ImpTotal").TEXT)
                ImporteNeto# = XVALO(node.selectSingleNode("ImpNeto").TEXT)
                ImporteNoGravado# = XVALO(node.selectSingleNode("ImpTotConc").TEXT)
                ImporteIva# = XVALO(node.selectSingleNode("ImpIVA").TEXT)
                ImporteExento# = XVALO(node.selectSingleNode("ImpOpEx").TEXT)
                ImportePercepcionIB# = XVALO(node.selectSingleNode("ImpTrib").TEXT)
                Moneda$ = node.selectSingleNode("MonId").TEXT
                TipoCambio# = XVALO(node.selectSingleNode("MonCotiz").TEXT)
                '
                Cae$ = node.selectSingleNode("CodAutorizacion").TEXT
                FechaVencimiento% = OPVARFA07.INVERFECHA%(node.selectSingleNode("FchVto").TEXT)
            Next node
            
            Set nodeListIIBB = doc.selectNodes("/FECompConsultaResponse/ResultGet/Tributos/Tributo")
            For Each nodeh In nodeListIIBB
                Jurisdiccion$ = nodeh.selectSingleNode("Desc").TEXT
                
                If InStr(1, Jurisdiccion$, "CABA", vbTextCompare) > 0 Then
                    BasePercepcionIBCaba# = XVALO(nodeh.selectSingleNode("BaseImp").TEXT)
                    PercepcionIBCaba# = XVALO(nodeh.selectSingleNode("Importe").TEXT)
                End If
                
                If InStr(1, Jurisdiccion$, "BSAS", vbTextCompare) > 0 Then
                    BasePercepcionIBBsAs# = XVALO(nodeh.selectSingleNode("BaseImp").TEXT)
                    PercepcionIBBsAs# = XVALO(nodeh.selectSingleNode("Importe").TEXT)
                End If
                
                If InStr(1, Jurisdiccion$, "MISIONES", vbTextCompare) > 0 Then
                    BasePercepcionIBMisiones# = XVALO(nodeh.selectSingleNode("BaseImp").TEXT)
                    PercepcionIBMisiones# = XVALO(nodeh.selectSingleNode("Importe").TEXT)
                End If
            Next nodeh
            
            Set nodeListIva = doc.selectNodes("/FECompConsultaResponse/ResultGet/Iva/AlicIva")
            For Each node In nodeListIva
                IdIva% = XVALO(node.selectSingleNode("Id").TEXT)
                If IdIva% = 1 Then
                    BaseImponibleNoGravada# = BaseImponibleNoGravada# + XVALO(node.selectSingleNode("BaseImp").TEXT)
                End If
                
                If IdIva% = 2 Then
                    BaseImponibleExenta# = BaseImponibleNoGravada# + XVALO(node.selectSingleNode("BaseImp").TEXT)
                End If
                
                If IdIva% = 3 Then
                    IvaTasa0# = 0
                End If
                
                If IdIva% = 4 Then
                    BaseImponible105# = BaseImponible105# + XVALO(node.selectSingleNode("BaseImp").TEXT)
                    IvaTasa105# = IvaTasa105# + XVALO(node.selectSingleNode("Importe").TEXT)
                End If
                
                If IdIva% = 5 Then
                    BaseImponible210# = BaseImponible210# + XVALO(node.selectSingleNode("BaseImp").TEXT)
                    IvaTasa210# = IvaTasa210# + XVALO(node.selectSingleNode("Importe").TEXT)
                End If
                
                If IdIva% = 6 Then
                    BaseImponible270# = BaseImponible270# + XVALO(node.selectSingleNode("BaseImp").TEXT)
                    IvaTasa270# = IvaTasa270# + XVALO(node.selectSingleNode("Importe").TEXT)
                End If
            Next
            
            ' CONTROLES
            
            'SumaBasesImponibles# = BaseImponibleNoGravada# + BaseImponibleExenta# + BaseImponible105# + BaseImponible210# + BaseImponible270#
            'If Abs(ImporteNeto# - SumaBasesImponibles#) > 0.005 Then
            '    Call MENSERR(24, "Imposible Recuperar Comprobante.\ \La Suma de la Base de Cada Item No Coincide\con el Importe Neto Total")
            '    MsgBox CStr(SumaBasesImponibles#)
            '    GoTo 99
            'End If
            
            'If Abs(ImporteIva# - (IvaTasa105# + IvaTasa210# + IvaTasa270#)) > 0.005 Then
            '    Call MENSERR(24, "Imposible Recuperar Comprobante.\ \Error al Obtener Discriminación del Iva")
            '    GoTo 99
            'End If
            
         End If
    End If
    '
    '********************************
    'PUNTO DE VENTA BIENES DE CAPITAL
    '********************************
    
    If PuntoVenta% = XVALO(CONTROL("FACTURA", "PUVENCAP")) Then
         Dim BC_doc As New MSXML2.DOMDocument
         Dim BC_success As Boolean

         BC_success = BC_doc.Load(CARPETOK$ + "CAE_" + Comprobante$ + ".TXT")

         Set nodeListEncabezado = BC_doc.selectNodes("/BFEGetCMPResponse/BFEResultGet")
         If Not nodeListEncabezado Is Nothing Then
            Dim BC_node As MSXML2.IXMLDOMNode
            Dim BC_nodeh As MSXML2.IXMLDOMNode

            For Each BC_node In nodeListEncabezado
                IDECAE$ = BC_node.selectSingleNode("Id").TEXT
                NumeroCliente& = OPVARFA07.CLIENTCUIT&(BC_node.selectSingleNode("Nro_doc").TEXT)
                FechaEmision% = OPVARFA07.INVERFECHA%(BC_node.selectSingleNode("Fecha_cbte_orig").TEXT)
                TotalComprobante# = XVALO(BC_node.selectSingleNode("Imp_total").TEXT)
                ImporteNeto# = XVALO(BC_node.selectSingleNode("Imp_neto").TEXT)
                ImporteNoGravado# = XVALO(BC_node.selectSingleNode("Imp_tot_conc").TEXT)
                ImporteIva# = XVALO(BC_node.selectSingleNode("Impto_liq").TEXT)
                ImporteExento# = XVALO(BC_node.selectSingleNode("Imp_op_ex").TEXT)
                ImportePercepcionIB# = XVALO(BC_node.selectSingleNode("Imp_iibb").TEXT)
                Moneda$ = BC_node.selectSingleNode("Imp_moneda_Id").TEXT
                TipoCambio# = XVALO(BC_node.selectSingleNode("Imp_moneda_ctz").TEXT)
                '
                Cae$ = BC_node.selectSingleNode("Cae").TEXT
                FechaVencimiento% = OPVARFA07.INVERFECHA%(BC_node.selectSingleNode("Fch_venc_Cae").TEXT)
            Next BC_node
            
            Dim Detalle() As String
            Dim CantidadItemsDetalle As Integer
            Set nodeListIIBB = BC_doc.selectNodes("/BFEGetCMPResponse/BFEResultGet/Items/Item")
            For Each BC_nodeh In nodeListIIBB
                CantidadItemsDetalle = CantidadItemsDetalle + 1
                ReDim Preserve Detalle(CantidadItemsDetalle)
                
                CodigoSecundario$ = BC_nodeh.selectSingleNode("Pro_codigo_sec").TEXT
                CodigoSecundarioInterno% = CODINT%(CodigoSecundario$)
                CodigoInterno% = 0
                If CodigoSecundarioInterno% > 0 Then
                    CodigoInterno% = CodigoSecundarioInterno%
                End If
                Detalle(CantidadItemsDetalle) = CStr(CodigoInterno%) & "|"
                
                Cantidad$ = BC_nodeh.selectSingleNode("Pro_qty").TEXT
                Detalle(CantidadItemsDetalle) = Detalle(CantidadItemsDetalle) & Cantidad$ & "|"
                
                PRECIO$ = BC_nodeh.selectSingleNode("Pro_precio_uni").TEXT
                Detalle(CantidadItemsDetalle) = Detalle(CantidadItemsDetalle) & PRECIO$ & "|"
                
                TotalItem$ = BC_nodeh.selectSingleNode("Imp_total").TEXT
                Detalle(CantidadItemsDetalle) = Detalle(CantidadItemsDetalle) & TotalItem$ & "|"

                CodigoIva$ = BC_nodeh.selectSingleNode("Iva_id").TEXT
                If XVALO(CodigoIva$) = 1 Then
                    BaseImponibleNoGravada# = BaseImponibleNoGravada# + XVALO(TotalItem$)
                End If
                
                If XVALO(CodigoIva$) = 2 Then
                    BaseImponibleExenta# = BaseImponibleExenta# + XVALO(TotalItem$)
                End If
                                
                If XVALO(CodigoIva$) = 4 Then
                    BaseImponible105# = BaseImponible105# + XVALO(TotalItem$)
                    IvaTasa105# = ROUND#(BaseImponible105# * 0.105)
                End If
                
                If XVALO(CodigoIva$) = 5 Then
                    BaseImponible210# = BaseImponible210# + XVALO(TotalItem$)
                    IvaTasa210# = ROUND#(BaseImponible210# * 0.21)
                End If
                
                If XVALO(CodigoIva$) = 6 Then
                    BaseImponible270# = BaseImponible270# + XVALO(TotalItem$)
                    IvaTasa270# = ROUND#(BaseImponible270# * 0.27)
                End If
                
            Next BC_nodeh
            
            If ImportePercepcionIB# > 0.005 Then
                CUEIVA$ = CONTROL$("PERIBCA", "CUECONTA")
                If CUECOINT%(CUEIVA$) > 0 Then
                    BasePercepcionIBCaba# = ImporteNeto# + ImporteNoGravado# + ImporteExento#
                    PercepcionIBCaba# = ImportePercepcionIB#
                Else
                    CUEIVA$ = CONTROL$("PERIBRU", "CUECONTA")
                    If CUECOINT%(CUEIVA$) > 0 Then
                        BasePercepcionIBBsAs# = ImporteNeto# + ImporteNoGravado# + ImporteExento#
                        PercepcionIBBsAs# = ImportePercepcionIB#
                    Else
                        CUEIVA$ = CONTROL$("PERIBMI", "CUECONTA")
                        If CUECOINT%(CUEIVA$) > 0 Then
                            BasePercepcionIBMisiones# = ImporteNeto# + ImporteNoGravado# + ImporteExento#
                            PercepcionIBMisiones# = ImportePercepcionIB#
                        End If
                    End If
                End If
            End If
            
            ' CONTROLES
            
            SumaBasesImponibles# = BaseImponibleNoGravada# + BaseImponibleExenta# + BaseImponible105# + BaseImponible210# + BaseImponible270#
            If Abs(ImporteNeto# - SumaBasesImponibles#) > 0.005 Then
                Call MENSERR(24, "Imposible Recuperar Comprobante.\ \La Suma de la Base de Cada Item No Coincide\con el Importe Neto Total")
                GoTo 99
            End If
            
            If Abs(ImporteIva# - (IvaTasa105# + IvaTasa210# + IvaTasa270#)) > 0.005 Then
                Call MENSERR(24, "Imposible Recuperar Comprobante.\ \Error al Obtener Discriminación del Iva")
                GoTo 99
            End If
            
         End If
         
    End If
    '
   
    '**************************
    'PUNTO DE VENTA EXPORTACION
    '**************************
    
    If PuntoVenta% = XVALO(CONTROL("FACTURA", "PUVEFAEX")) Then
         Dim EXPO_doc As New MSXML2.DOMDocument
         Dim EXPO_success As Boolean

         BC_success = EXPO_doc.Load(CARPETOK$ + "CAE_" + Comprobante$ + ".TXT")

         Set nodeListEncabezado = EXPO_doc.selectNodes("/FEXGetCMPResponse/FEXResultGet")
         If Not nodeListEncabezado Is Nothing Then
            Dim EXPO_node As MSXML2.IXMLDOMNode
            Dim EXPO_nodeh As MSXML2.IXMLDOMNode

            For Each EXPO_node In nodeListEncabezado
                IDECAE$ = EXPO_node.selectSingleNode("Id").TEXT
                CuitCliente$ = EXPO_node.selectSingleNode("Cuit_pais_cliente").TEXT
                If XVALO(Mid$(CuitCliente$, 1, 2)) > 50 Then
                    CuentaCliente$ = EXPO_node.selectSingleNode("Obs").TEXT
                    N0% = InStr(1, CuentaCliente$, "(")
                    N1% = InStr(1, CuentaCliente$, ")")
                    If N0% > 0 And N1% And N1% > N0% Then
                       NumeroCliente& = XVALO(Mid$(CuentaCliente$, N0% + 1, N1% - N0% - 1))
                    End If
                Else
                    NumeroCliente& = OPVARFA07.CLIENTCUIT&(CuitCliente$)
                End If
                
                FechaEmision% = OPVARFA07.INVERFECHA%(EXPO_node.selectSingleNode("Fecha_cbte").TEXT)
                TotalComprobante# = XVALO(EXPO_node.selectSingleNode("Imp_total").TEXT)
                ImporteNeto# = TotalComprobante#
                ImporteNoGravado# = 0
                ImporteIva# = 0
                ImporteExento# = 0
                ImportePercepcionIB# = 0
                Moneda$ = EXPO_node.selectSingleNode("Moneda_Id").TEXT
                TipoCambio# = XVALO(EXPO_node.selectSingleNode("Moneda_ctz").TEXT)
                '
                Cae$ = EXPO_node.selectSingleNode("Cae").TEXT
                FechaVencimiento% = OPVARFA07.INVERFECHA%(EXPO_node.selectSingleNode("Fch_venc_Cae").TEXT)
            Next EXPO_node
            
            Dim DetalleExpo() As String
            Dim CantidadItemsDetalleExpo As Integer
            Set nodeListIIBB = EXPO_doc.selectNodes("/FEXGetCMPResponse/FEXResultGet/Items/Item")
            For Each EXPO_nodeh In nodeListIIBB
                CantidadItemsDetalleExpo = CantidadItemsDetalleExpo + 1
                ReDim Preserve DetalleExpo(CantidadItemsDetalleExpo)
                
                On Error Resume Next
                Codigo$ = EXPO_nodeh.selectSingleNode("Pro_codigo").TEXT
                If Err > 0 Then
                    ReDim Preserve DetalleExpo(CantidadItemsDetalleExpo - 1)
                    GoTo 90
                End If
                On Error GoTo 0
                
                CodigoInterno% = CODINT%(Codigo$)
                DetalleExpo(CantidadItemsDetalleExpo) = CStr(CodigoInterno%) & "|"
                    
                Cantidad$ = EXPO_nodeh.selectSingleNode("Pro_qty").TEXT
                DetalleExpo(CantidadItemsDetalleExpo) = DetalleExpo(CantidadItemsDetalleExpo) & Cantidad$ & "|"
                
                PRECIO$ = EXPO_nodeh.selectSingleNode("Pro_precio_uni").TEXT
                DetalleExpo(CantidadItemsDetalleExpo) = DetalleExpo(CantidadItemsDetalleExpo) & PRECIO$ & "|"
                
                TotalItem$ = EXPO_nodeh.selectSingleNode("Pro_total_item").TEXT
                DetalleExpo(CantidadItemsDetalleExpo) = DetalleExpo(CantidadItemsDetalleExpo) & TotalItem$ & "|"

                SumaBasesImponibles# = SumaBasesImponibles# + XVALO(TotalItem$)
90          Next EXPO_nodeh
                        
            ' CONTROLES
            
            If Abs(TotalComprobante# - SumaBasesImponibles#) > 0.005 Then
                Call MENSERR(24, "Imposible Recuperar Comprobante.\ \La Suma de la Base de Cada Item No Coincide\con el Importe Total")
                GoTo 99
            End If
         End If
    End If
    
    Close NARCHIVO%
    
    'GRABACION EN EL CAJABANC
    '************************
    '
    ' OBTENCION DE LOS DATOS
    NFCY$ = Comprobante$
    NFCX$ = Left$(Comprobante$, 3) + Mid$(Comprobante$, 11, 6) + Right(Comprobante$, 2)
    OPMOVIX$ = UCase(Left$(Comprobante$, 2))
    LETRA$ = Right(NFCY$, 1)
    PTOVTASTR$ = Mid$(NFCY$, 4, 4)
    SIGX% = 1
    If OPMOVIX$ = "NC" Then SIGX% = (-1)  ' SI ES NOTA DE CREDITO
    
    ' COMPROBANTE EN MONEDA EXTRANJERA
    If TipoCambio# < 0.00005 Then TipoCambio# = 1
    MONEMI% = OPVARFA07.TIPOMONEDA%(Moneda$)
    COTIZADOLAR# = TICAMONE(2): If COTIZADOLAR# < 0.0005 Then COTIZADOLAR# = 1
    
    ' GRABA MOVIMIENTO FVEN
    CoComCor$ = NFCX$
    CoComLar$ = NFCY$
    NumeClie = Int(GRUDIVCLI(NumeroCliente&))
    NCliReal = NumeroCliente&
    NumeProv% = 0
    FechEmi = CVDAT(FechaEmision%)
    Referen$ = rasocli$(NumeroCliente&)
    FeContab = CVDAT(FechaEmision%)
    FeAcredi = CVDAT(FechaEmision%)
    CuentaCo% = CUECOCLI%(NumeroCliente&)
    TiMovi$ = "FVEN"
    OpMovi$ = OPMOVIX$
    VaMovi$ = LETRA$
    DOCUMEN$ = "Factura:": If OPMOVIX$ = "NC" Then DOCUMEN$ = "N.Crédito:": If OPMOVIX$ = "ND" Then DOCUMEN$ = "N.Débito:"
    Descrip$ = DOCUMEN$ + " " + LETRA$ + " (Mat.Codif.)" + " " + PTOVTASTR$ + "-" + TRIM$(Str(XVALO(Mid$(NFCY$, 9, 8))))
    NumeTal = XVALO(PTOVTASTR$)
    NumeCom = XVALO(Mid$(NFCY$, 9, 8))
    StatInit = 0
    ImpoDeb = TotalComprobante# * SIGX% * TipoCambio#
    ValoDola = TotalComprobante# * SIGX%
    Observa$ = rasocli$(NumeroCliente&)
    NumeroCae$ = Cae$
    IdentiCae$ = IDECAE$
    FechVtoCae% = FechaVencimiento%
    ObservaComprobante$ = "Factura Recuperada Automaticamente Desde AFIP"
    COTIMONEM# = TipoCambio#
    Call GRACABAN
    '
    ' GRABA MOVIMIENTO VENTA
    If ImporteNeto# + ImporteNoGravado# >= 0.005 Then
      PPVI% = PIVACLI%(NumeroCliente&)
      If PPVI% < 0 Or PPVI% > 7 Then PPVI% = 1
      CUENTAX$ = CONTROL$("SVD451S", "CUEVE-" + TRIM$(Str$(PPVI%)))
      QQ1% = CUECOINT%(CUENTAX$)
      CuentaCo% = QQ1%
      TiMovi$ = "VENTA"
      OpMovi$ = OPMOVIX$
      VaMovi$ = LETRA$
      Descrip$ = "Mercaderia Vendida " + NFCX$
      StatInit = 0
      ImpoDeb = (ImporteNeto# + ImporteNoGravado#) * SIGX% * (-1) * TipoCambio#
      ValoDola = ImporteNeto# + ImporteNoGravado#
      Call GRACABAN
    End If
    '
    ' GRABA MOVIMIENTO IVA DEBITO
    CUEIVX$(0) = CONTROL$("SVD451S", "CUEIVA") ' 10.5%
    CUEIVX$(1) = CONTROL$("SVD451S", "CUEIVB") ' 21.0%
    If ImporteIva# >= 0.005 Then
       If IvaTasa105# > 0.005 Then
            CuentaCo% = CUECOINT%(CUEIVX$(0)) 'SE IMPUTA A LA CUENTA CONTABLE DE IVA PRINCIPAL
            TiMovi$ = "IVA"
            OpMovi$ = "DeFi"
            VaMovi$ = "10.50 %"
            Descrip$ = "Iva Debito Fiscal " + NFCX$
            StatInit = 0
            ImpoDeb = IvaTasa105# * SIGX% * (-1) * TipoCambio#
            ValoDola = IvaTasa105# * SIGX% * (-1)
            BasImpo# = BaseImponible105# * (-1) * TipoCambio#
            Call GRACABAN
       End If
       
       If IvaTasa210# > 0.005 Then
            CuentaCo% = CUECOINT%(CUEIVX$(1)) 'SE IMPUTA A LA CUENTA CONTABLE DE IVA PRINCIPAL
            TiMovi$ = "IVA"
            OpMovi$ = "DeFi"
            VaMovi$ = "21.00 %"
            Descrip$ = "Iva Debito Fiscal " + NFCX$
            StatInit = 0
            ImpoDeb = IvaTasa210# * SIGX% * (-1) * TipoCambio#
            ValoDola = IvaTasa210# * SIGX% * (-1)
            BasImpo# = BaseImponible210# * (-1) * TipoCambio#
            Call GRACABAN
       End If
       
       If IvaTasa270# > 0.005 Then
            CuentaCo% = CUECOINT%(CUEIVX$(3)) 'SE IMPUTA A LA CUENTA CONTABLE DE IVA PRINCIPAL
            TiMovi$ = "IVA"
            OpMovi$ = "DeFi"
            VaMovi$ = "27.00 %"
            Descrip$ = "Iva Debito Fiscal " + NFCX$
            StatInit = 0
            ImpoDeb = IvaTasa270# * SIGX% * (-1) * TipoCambio#
            ValoDola = IvaTasa270# * SIGX% * (-1)
            BasImpo# = BaseImponible270# * (-1) * TipoCambio#
            Call GRACABAN
       End If
       '
   End If
   '
   ' GRABA MOVIMIENTO PERCEPCION IIBB
   If ImportePercepcionIB# >= 0.005 Then
        
        If PercepcionIBCaba# > 0.005 Then
            CUEIVA$ = CONTROL$("PERIBCA", "CUECONTA")
            CuentaCo% = CUECOINT%(CUEIVA$)
            TiMovi$ = "IBRU"
            OpMovi$ = "IB-C"
            VaMovi$ = "PERC"
            Descrip$ = "Perc.IB. CABA " + NFCX$
            StatInit = 0
            ImpoDeb = PercepcionIBCaba# * SIGX% * (-1) * TipoCambio#
            ValoDola = PercepcionIBCaba# * SIGX% * (-1)
            BasImpo# = BasePercepcionIBCaba#
            Call GRACABAN
        End If
        
        If PercepcionIBBsAs# > 0.005 Then
            CUEIVA$ = CONTROL$("PERIBRU", "CUECONTA")
            CuentaCo% = CUECOINT%(CUEIVA$)
            TiMovi$ = "IBRU"
            OpMovi$ = "IB-B"
            VaMovi$ = "PERC"
            Descrip$ = "Perc.IB.Bs.As. " + NFCX$
            StatInit = 0
            ImpoDeb = PercepcionIBBsAs# * SIGX% * (-1) * TipoCambio#
            ValoDola = PercepcionIBBsAs# * SIGX% * (-1)
            BasImpo# = BasePercepcionIBBsAs#
            Call GRACABAN
        End If
        
        If PercepcionIBMisiones# > 0.005 Then
            CUEIVA$ = CONTROL$("PERIBMI", "CUECONTA")
            CuentaCo% = CUECOINT%(CUEIVA$)
            TiMovi$ = "IBRU"
            OpMovi$ = "IB-M"
            VaMovi$ = "PERC"
            Descrip$ = "Perc.IB. MISIONES " + NFCX$
            StatInit = 0
            ImpoDeb = PercepcionIBMisiones# * SIGX% * (-1) * TipoCambio#
            ValoDola = PercepcionIBMisiones# * SIGX%
            BasImpo# = BasePercepcionIBMisiones#
            Call GRACABAN
        End If
   
   End If
   BasImpo# = 0
   '
   '
   'GRABA MOVIMIENTO DE STOCK
   '*************************
   '
   If PuntoVenta% = XVALO(CONTROL("FACTURA", "PUVENELE")) Or PuntoVenta% = XVALO(CONTROL("FACTURA", "PUVECRED")) Then      ' FACTURA COMUN
        
        CANTIDAD0# = (-1)
        If Mid$(NFCX$, 1, 2) = "NC" Then CANTIDAD0# = 1
        ImporteUnitario# = ImporteNeto# + ImporteNoGravado# + ImporteExento#
        If Mid$(NFCX$, 11, 1) = "B" Then ImporteUnitario# = TotalComprobante#
        Call MOVISTOK(FechaEmision%, 0, "", "RT", NFCY$, NFCX$, NFCY$, NFCX$, "Recuperacion desde AFIP", 0, ImporteUnitario#, MONEMI%, NumeroCliente&, 0, CANTIDAD0#, 0, 1, "Recuperacion desde AFIP", "", "", ImporteUnitario#, 0, "", 0)
        
   ElseIf PuntoVenta% = XVALO(CONTROL("FACTURA", "PUVENCAP")) Then  ' FACTURA BS DE CAPITAL
        
        For ItemDetalle% = 1 To UBound(Detalle)
            RegistroDetalle$ = Detalle(ItemDetalle%)
                INI% = 1
                FIN% = InStr(INI%, RegistroDetalle, "|")
            CodigoInterno% = XVALO(Mid$(RegistroDetalle$, INI%, FIN% - 1))
                INI% = FIN% + 1
                FIN% = InStr(INI%, RegistroDetalle, "|")
            CANTIDAD0# = (-1) * XVALO(Mid$(RegistroDetalle$, INI%, FIN% - 1))
            If Mid$(NFCX$, 1, 2) = "NC" Then CANTIDAD0# = 1
                INI% = FIN% + 1
                FIN% = InStr(INI%, RegistroDetalle, "|")
            ImporteNeto# = XVALO(Mid$(RegistroDetalle$, INI%, FIN% - 1))
                INI% = FIN% + 1
                FIN% = InStr(INI%, RegistroDetalle, "|")
            ImporteItem# = XVALO(Mid$(RegistroDetalle$, INI%, FIN% - 1))
            
            If Mid$(NFCX$, 18, 1) = "B" Then ImporteUnitario# = TotalComprobante#
            
            Call MOVISTOK(FechaEmision%, CodigoInterno%, "", "RT", NFCY$, NFCX$, NFCY$, NFCX$, "Recuperacion desde AFIP", ItemDetalle% - 1, ImporteNeto#, MONEMI%, NumeroCliente&, 0, CANTIDAD0#, 0, 0, "Recuperacion desde AFIP", "", "", ImporteNeto#, 0, "", 0)
        Next
        
   ElseIf PuntoVenta% = XVALO(CONTROL("FACTURA", "PUVEFAEX")) Then  ' FACTURA EXPORTACION
   
        For ItemDetalle% = 1 To UBound(DetalleExpo)
            RegistroDetalle$ = DetalleExpo(ItemDetalle%)
                INI% = 1
                FIN% = InStr(INI%, RegistroDetalle, "|")
            CodigoInterno% = XVALO(Mid$(RegistroDetalle$, INI%, FIN% - 1))
                INI% = FIN% + 1
                FIN% = InStr(INI%, RegistroDetalle, "|")
            CANTIDAD0# = (-1) * XVALO(Mid$(RegistroDetalle$, INI%, FIN% - 1))
            If Mid$(NFCX$, 1, 2) = "NC" Then CANTIDAD0# = 1
                INI% = FIN% + 1
                FIN% = InStr(INI%, RegistroDetalle, "|")
            ImporteNeto# = XVALO(Mid$(RegistroDetalle$, INI%, FIN% - 1))
            ImporteNeto# = ImporteNeto# / 10 ' ESTO ES PARA CORREGIR UN ERROR EN EL .NET QUE PASA EL PRECIO DEL ITEM COMO PRECIO TOTAL
                INI% = FIN% + 1
                FIN% = InStr(INI%, RegistroDetalle, "|")
            ImporteItem# = XVALO(Mid$(RegistroDetalle$, INI%, FIN% - 1))
            
            Call MOVISTOK(FechaEmision%, CodigoInterno%, "", "RT", NFCY$, NFCX$, NFCY$, NFCX$, "Recuperacion desde AFIP", ItemDetalle% - 1, ImporteNeto#, MONEMI%, NumeroCliente&, 0, CANTIDAD0#, 0, 0, "Recuperacion desde AFIP", "", "", ImporteNeto#, 0, "", 0)
        Next
   
   End If
   '
   '
   'GRABA LA DEUDA EN SAUDEPEN
   '**************************
   '
   COPAGI% = CPAGCLI%(NC)
   VARIDOC% = 1
   '
   If LETRA$ = "B" Then VARIDOC% = 2
   If LETRA$ = "E" Then VARIDOC% = 3
   Tipo% = 2
   If OPMOVIX$ = "ND" Then Tipo% = 3
   If OPMOVIX$ = "NC" Then Tipo% = 6
   '
   Call FORFASQL.GRADEUFA(Tipo%, VARIDOC%, NumeCom, NumeroCliente&, FechaEmision%, COPAGI%, TotalComprobante# * TipoCambio#)
   '
'   'GRABAR COMPROBANTE EN LISTA DE FACTURAS RECUPERADAS
'   '***************************************************
'   '
'10 RNDX$ = RANDSTRING$(-10)
'   If CantRegistros("FARECAFIP", "NFACTURA = '" & RNDX$ & "'", "NOMESS") > 0 Then GoTo 10
'   Call CreaRegistro("FARECAFIP", "NFACTURA", RNDX$, "NOMESS")
'   Condi$ = "NFACTURA = '" & RNDX$ & "'"
'   Call ACTUREGISTRO("FARECAFIP", "FECHEMIS", Condi$, FETEXCOR1$(FechaEmision%), REGAF&, "NOMESS")
'   Call ACTUREGISTRO("FARECAFIP", "IMPORTE", Condi$, TotalComprobante#, REGAF&, "NOMESS")
'   Call ACTUREGISTRO("FARECAFIP", "NuClien", Condi$, NumeroCliente&, REGAF&, "NOMESS")
'   Call ACTUREGISTRO("FARECAFIP", "NFACTURA", Condi$, NFCY$, REGAF&, "NOMESS")
'   '
99 End Sub
Function BuscaComprobanteAfipEnSistema(ComprobanteARecuperar$, FechaEmision$, CuitCliente$, ImporteTotal#, Cae$, VencimientoCae$, NroCliente&) As Boolean
   
   
    BuscaComprobanteAfipEnSistema = True
    
    ' OBTENGO CLIENTE ASOCIADO AL COMPROBANTE A RECUPERAR
    Cuit$ = Mid$(CuitCliente$, 1, 2) & "-" & Mid$(CuitCliente$, 3, 8) & "-" & Mid$(CuitCliente$, 11, 1)
    If Len(CuitCliente$) < 11 Then
        Cuit$ = Mid$(CuitCliente$, 1, 2) & "." & Mid$(CuitCliente$, 3, 3) & "." & Mid$(CuitCliente$, 6, 3)
    End If
    ConsultaSql$ = "select nume_cli, raso_cli from deudor12 with(nolock) where (cuit_cli='" & Cuit$ & "' or cuit_cli='" & CuitCliente$ & "') and stat_cli>0"
    ' OJO !!! ACA EXISTE LA POSIBILIDAD DE MAS DE UN CLIENTE CON EL MISMO CUIT
    ' CON GRU_DIV=0 SE ESTA ELIMINANDO A LAS SUCURSALES.
    Dim rst_Clientes As ADODB.Recordset
    Set rst_Clientes = READSET(ConsultaSql$)
    CantidadDeCoincidencias& = rst_Clientes.RecordCount
    
    If CantidadDeCoincidencias& > 0 Then
        
        ' EN CASO DE QUE EXISTA MAS DE UN CLIENTE CON EL MISMO CUIT, DEJA AL USUARIO
        ' LA SELECCION DEL NUMERO DE CUENTA AL QUE PERTENECE EL COMPROBANTE A RECUPERAR
        If CantidadDeCoincidencias& > 1 Then
                'CAMPOS$ = "Cuenta/I10;Cliente/A48"
                'Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, CAMPOS$, FRMZELECHO)
                'Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, CAMPOS$, FRMZELECHO)
                Call COPYSTRU("DEUDOR12", "CLIREPE")
                
                J& = 0
                With rst_Clientes
                    Do While Not .EOF
                        J& = J& + 1
                        'FRMZELECHO.MSF2.Rows = J& + 1
                        'FRMZELECHO.MSF2.TextMatrix(J&, 1) = CStr(!nume_cli)
                        'FRMZELECHO.MSF2.TextMatrix(J&, 2) = CStr(!raso_cli)
                        XX$ = REGBLAN$("CLIREPE")
                        Call REPLA(XX$, MKS$(!Nume_Cli), 1, 4)
                        Call REPLA(XX$, !raso_cli, 5, 64)
                        
                        Call GRAREG("CLIREPE", XX$, J&)
                        
                        .MoveNext
                    Loop
                End With
                
                Call SETULTREG("CLIREPE", J&)
                'FRMZELECHO.Width = 7000
                Screen.MousePointer = 1
                'FRMZELECHO.Show 1
                Call SELECHO("CLIREPE")
                
                'NumeroCliente& = XVALO(RESP_ZELE_VECT(1))
                TX$ = VALACT1$("CLIREPE")
                NumeroCliente& = XVALO(TX$)
                
                If NumeroCliente& = 0 Then GoTo 99
        Else
            NumeroCliente& = rst_Clientes!Nume_Cli
        End If
        NroCliente& = NumeroCliente&
        Screen.MousePointer = 11
        On Error Resume Next
        rst_Clientes.Close
        Set rst_Clientes = Nothing
        On Error GoTo 0
    
        ' BUSCA EL COMPROBANTE SIN CAE (NUMECOMP<0) QUE COINCIDE EN NRO DE CLIENTE, PUNTO DE VENTA E IMPORTE
        CodigoCbte$ = Mid$(ComprobanteARecuperar$, 1, 2)
        LETRA$ = Mid$(ComprobanteARecuperar$, 18, 1)
        PuntoVenta$ = CStr(XVALO(Mid$(ComprobanteARecuperar$, 4, 4)))
        NumeroComprobante$ = CStr(XVALO(Mid$(ComprobanteARecuperar$, 9, 8)))
        CondicionSql$ = "numecomp<0 and tipomovim='fven' " _
        & " and cli_real=" & CStr(NumeroCliente&) _
        & " and opcimovim='" & Mid$(CodigoCbte$, 1, 2) & "'" _
        & " and varimovim='" & LETRA$ & "'" _
        & " and numetalo=" & PuntoVenta$ _
        & " and abs(impo_dola-" & TRIM$(FORMATNUM$(ImporteTotal#, "F12.2")) & ")<0.05"
        IdTransaccion& = XVALO(VALOBADA("cajabanc", "idcaenum", CondicionSql$))
        
        If IdTransaccion& > 0 Then
            Text5 = CStr(IdTransaccion&)
            Label6 = VencimientoCae$
            Text1 = Cae$
                    
            Screen.MousePointer = 1
            
            ModoRecuperacionDeComprobante = True
            
            Text2 = NumeroComprobante$
            
            If XVALO(Text1) > 0 Then
               PREFI$ = "0000" & TRIM$(Text4): PREFI$ = Right(PREFI$, 4)
               NumeroFactura$ = "00000000" & TRIM$(Text2): NumeroFactura$ = Right(NumeroFactura$, 8)
               NUMFACTUNUE$ = Mid$(Text6, 1, 3) & PREFI$ + "-" + NumeroFactura & Mid$(Text6, 17, 2)
               NUFACTUVIE$ = TRIM$(Text6)
               lblFechaEmisionAfip = FechaEmision$
               
               If XVALO(Text5) > 0 And NUFACTUVIE$ <> "" Then
                 Call GRABANUEFACTU(NUFACTUVIE$, NUMFACTUNUE$)
               Else
                 Call COMUNI("Imposible Continuar.\Numero de Id Erroneo o Falta de Numero Provisorio.\Favor Comunicarse Con Mesa de Ayuda Flexoft")
               End If
            End If
            
            ModoRecuperacionDeComprobante = False
            
        Else
            Screen.MousePointer = 1
            BuscaComprobanteAfipEnSistema = False
        End If
    
    Else
        Screen.MousePointer = 1
        Call COMUNI("Imposible Recuperar.\No se Ha Encontrado el Cliente Asociado al Cuit '" & CuitCliente$ & "'.")
        
        BuscaComprobanteAfipEnSistema = False
    End If
     
    Exit Function
    
99  Screen.MousePointer = 1
    BuscaComprobanteAfipEnSistema = False
    
End Function

Private Sub tmrRecuperaUltimoNroCbte_Timer()

    tmrRecuperaUltimoNroCbte.Enabled = False
    Screen.MousePointer = 11
    
    Static CantidadTiposProcesados As Integer
    Static EnProceso As Boolean
    Static VUELTA As Integer
    Static ArchivoUltimoNumero As String
    Static PuntoVentaComun As Long
    Static PuntoVentaCredito As Long
    Static PuntoVentaBienes As Long
    Static PuntoVentaExportacion As Long
    Dim REGISTRO As String
    Dim UltimoNumero As Long
    
    If VUELTA = 0 Then
        ArchivoUltimoNumero = "C:\FLEXOFT\FACELEC\ULTNUM.TXT"
        CantidadTiposProcesados = 0
        EnProceso = False
        
        PuntoVentaComun = XVALO(CONTROL("FACTURA", "PUVENELE"))
        PuntoVentaBienes = XVALO(CONTROL("FACTURA", "PUVENCAP"))
        PuntoVentaExportacion = XVALO(CONTROL("FACTURA", "PUVEFAEX"))
        PuntoVentaCredito = XVALO(CONTROL("FACTURA", "PUVECRED"))
        If PuntoVentaCredito = 0 Then PuntoVentaCredito = PuntoVentaComun
        
        VUELTA = 1
        Text3 = "Recuperando último comprobante registrado ..."
        Text3.Refresh
        
        Call REMEVE07.BusquedaHuecos_StoreProcedure
    End If
    
    If Not EnProceso Then       ' NO ESTA EN PROCESO
        
        If CantidadTiposProcesados < UBound(RecuperacionTipoComprobante%) Then
            If EXISTE%(ArchivoUltimoNumero) = 1 Then
                Kill ArchivoUltimoNumero
            End If
            TipoCbte$ = RecuperacionTipoComprobante%(CantidadTiposProcesados)
            Select Case RecuperacionPuntoVenta%
                Case PuntoVentaComun, PuntoVentaCredito: r = Shell("c:\flexoft\facelec\wsfev1_cliente.exe ULTNUM " & TipoCbte$ & " " & CStr(RecuperacionPuntoVenta%), vbMinimizedNoFocus)
                
                Case PuntoVentaBienes: r = Shell("c:\flexoft\facelec\wsbfev1_cliente.exe ULTNUM " & TipoCbte$ & " " & CStr(RecuperacionPuntoVenta%), vbMinimizedNoFocus)
                
                Case PuntoVentaExportacion: r = Shell("c:\flexoft\facelec\wsfexv1_cliente.exe ULTNUM " & TipoCbte$ & " " & CStr(RecuperacionPuntoVenta%), vbMinimizedNoFocus)
            End Select
            
            EnProceso = True
            CantidadTiposProcesados = CantidadTiposProcesados + 1
        Else
            VUELTA = 0
            Screen.MousePointer = 1
            Call BuscaHuecosDeComprobantesEnSistema
        End If
        
    Else                        ' ESTA EN PROCESO
        
        If EXISTE%(ArchivoUltimoNumero) > 0 Then
            NARCHIVO% = FILEOPEN%(ArchivoUltimoNumero, 1, 128)  'ABRE EL ARCHIVO DE LECTURA
            Do While Not EOF(NARCHIVO%) ' RECORRE
                Line Input #NARCHIVO%, REGISTRO ' LEE CADA REGISTRO
                UltimoNumero = CLng(REGISTRO)
            Loop
            Close #NARCHIVO%
            
            RecuperacionUltimoNumeroPorTipo&(CantidadTiposProcesados - 1) = UltimoNumero
            
            EnProceso = False
            tmrRecuperaUltimoNroCbte.Enabled = True
        End If
    
    End If
    
    If EnProceso Then tmrRecuperaUltimoNroCbte.Enabled = True
    
End Sub


Sub BuscaHuecosDeComprobantesEnSistema()

    Dim LETRA$, Tipo$
    
    Screen.MousePointer = 11
    
    TiposProcesados% = 0
    On Error Resume Next
    CantidadHuecos = UBound(RecuperacionHuecos$)
    On Error GoTo 0
    If CantidadHuecos > 0 Then ReDim RecuperacionHuecos$(0)
    
    Call REMEVE07.AnalizaCorrelativaDeComprobantes(0, False)
    
    Do While TiposProcesados% < UBound(RecuperacionTipoComprobante%)
            
            EsModoCredito = False
            Select Case RecuperacionTipoComprobante%(TiposProcesados%)
                Case 1: Tipo$ = "FC": LETRA$ = "A"
                Case 2: Tipo$ = "ND": LETRA$ = "A":
                Case 3: Tipo$ = "NC": LETRA$ = "A"
                Case 6: Tipo$ = "FC": LETRA$ = "B"
                Case 7: Tipo$ = "ND": LETRA$ = "B"
                Case 8: Tipo$ = "NC": LETRA$ = "B"
                Case 19: Tipo$ = "FC": LETRA$ = "E"
                Case 20: Tipo$ = "ND": LETRA$ = "E"
                Case 21: Tipo$ = "NC": LETRA$ = "E"
                Case 201: Tipo$ = "FC": LETRA$ = "A"
                Case 202: Tipo$ = "ND": LETRA$ = "A"
                Case 203: Tipo$ = "NC": LETRA$ = "A"
            End Select
            
            ' BUSCO INICIO
            Dim RstInicio As ADODB.Recordset
            DESDE% = DIANTE(HOY(HO$), 75)
            ConsultaSql$ = "select max(numecomp) as iniciobusqueda from cajabanc with(nolock) where " _
            & " nuclien>0 and tipomovim='fven' and opcimovim='" & Tipo$ & "' and varimovim='" & LETRA$ & "'" _
            & " and numetalo=" & CStr(RecuperacionPuntoVenta%) & " and fechconta<='" & FETEXCOR1$(DESDE%) & "'"
            Set RstInicio = READSET(ConsultaSql$)
'            Dim RstInicio As ADODB.Recordset
'            DESDE% = FECHANUM("01/01/20")
'            ConsultaSql$ = "select max(numecomp) as iniciobusqueda from cajabanc with(nolock) where " _
'            & " nuclien>0 and tipomovim='fven' and opcimovim='" & Tipo$ & "' and varimovim='" & LETRA$ & "'" _
'            & " and numetalo=" & CStr(RecuperacionPuntoVenta%) & " and fechconta>'" & FETEXCOR1$(FECHANUM("01/06/19")) & "' and fechconta<'" & FETEXCOR1$(DESDE%) & "'"
'            Set RstInicio = READSET(ConsultaSql$)
            
            If XVALO(RstInicio!iniciobusqueda) > 0 Then
                INI = XVALO(RstInicio!iniciobusqueda) + 1
            Else
                INI = 1
            End If
            On Error Resume Next
            RstInicio.Close
            Set RstInicio = Nothing
            On Error GoTo 0
            
            ' BUSCO FIN
            FIN = RecuperacionUltimoNumeroPorTipo&(TiposProcesados%)
            
            Text3 = Text3 & vbCrLf & "Buscando Comprobante Faltantes ..."
            Text3 = Text3 & vbCrLf & Tipo$ & "-" & LETRA$ & ": " & CStr(INI) & " a " & CStr(FIN)
            J& = 0
            On Error Resume Next
            If UBound(RecuperacionHuecos$) > 0 Then J& = UBound(RecuperacionHuecos$)
            On Error GoTo 0
            
            'MsgBox Tipo$ & vbCrLf & LETRA$ & vbCrLf & CStr(RecuperacionPuntoVenta%) & vbCrLf & INI & vbCrLf & FIN
            Resultado$ = REMEVE07.BusquedaHuecos(Tipo$, LETRA$, CStr(RecuperacionPuntoVenta%), CLng(INI), CLng(FIN))
'            For i = INI To FIN
'                CondicionSql$ = "fechconta>='" & FETEXCOR1$(DESDE%) & "' and opcimovim='" & Tipo$ & "' and varimovim='" & LETRA$ & "' and numetalo=" & CStr(RecuperacionPuntoVenta%) & " and numecomp=" & CStr(i) & " and nuclien>0 and tipomovim='fven'"
'                If RecuperacionModoCredito = False Then
'                    CondicionSql$ = CondicionSql$ & " and (modoopcimovim='' or modoopcimovim is null)"
'                Else
'                    CondicionSql$ = CondicionSql$ & " and modoopcimovim='cr'"
'                End If
'                If CantRegistros("cajabanc", CondicionSql$) = 0 Then
'
'                   J& = J& + 1
'                   Comprobante$ = Tipo$ & "-" & Right$("0000" & CStr(RecuperacionPuntoVenta%), 4) & "-" & Right$("00000000" & CStr(i), 8) & "-" & LETRA$
'                   ReDim Preserve RecuperacionHuecos$(J&)
'
'                   RecuperacionHuecos$(J&) = Comprobante$
'                End If
'            Next i
            If Resultado$ <> "OK" And Resultado$ <> "" Then
                Call TEXTOLOTES$(Resultado$, ",")
                For iHuecos& = 1 To UBound(CADENATEX$)
                    J& = J& + 1
                    Comprobante$ = Tipo$ & "-" & Right$("0000" & CStr(RecuperacionPuntoVenta%), 4) & "-" & Right$("00000000" & CStr(CLng(CADENATEX$(iHuecos&))), 8) & "-" & LETRA$
                    ReDim Preserve RecuperacionHuecos$(J&)

                    RecuperacionHuecos$(J&) = Comprobante$
                Next iHuecos&
            End If
            
            TiposProcesados% = TiposProcesados% + 1
    Loop
       
    On Error Resume Next
    TotalComprobantesARecuperar& = 0
    TotalComprobantesARecuperar& = UBound(RecuperacionHuecos$)
    On Error GoTo 0
    
    If TotalComprobantesARecuperar& > 0 Then
        tmrRecuperaComprobanteDesdeAfip.Enabled = True
    Else
        If ModoSolicitudManual = True Then
            Text3 = ""
            If CantRegistros("cajabanc", "codicom_lar='" & Text6 & "' and tipomovim='fven'") > 0 Then
                Call SOLICITACAE
            End If
            ModoSolicitudManual = False
            Command4.Enabled = True
            Screen.MousePointer = 1
        Else
            ModoSolicitudManual = False
            Command4.Enabled = True
            Screen.MousePointer = 1
            Call COMUNI("No Se Han Encontrado Comprobantes A Recuperar")
        End If
    End If
    
End Sub


