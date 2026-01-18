VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Begin VB.Form BALANZA 
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lectura de Peso desde Balanza"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   6285
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtTest 
      Appearance      =   0  'Flat
      Height          =   975
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   5160
      Width           =   5775
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   120
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.TextBox txtLecturaContinua 
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
      Height          =   1320
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   3720
      Width           =   5775
   End
   Begin VB.Timer tmrObtienePesada 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   120
      Top             =   120
   End
   Begin VB.CommandButton btnRefresh 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      Picture         =   "BALANZA.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton btnCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   6
      Top             =   3000
      Width           =   2055
   End
   Begin VB.CommandButton btnAceptarPeso 
      Caption         =   "Aceptar Peso"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   5
      Top             =   3000
      Width           =   2055
   End
   Begin VB.TextBox txtPeso 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   1800
      TabIndex        =   3
      Top             =   1680
      Width           =   3135
   End
   Begin VB.Label lblIntentos 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   10
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label lblLectura 
      BackStyle       =   0  'Transparent
      Caption         =   "Leyendo balanza"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   8
      Top             =   2400
      Width           =   3015
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Peso"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label txtDescripcion 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   5775
   End
   Begin VB.Label txtCodigo 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   540
      Width           =   615
   End
End
Attribute VB_Name = "BALANZA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Respuesta As String
Public CodigoProducto As String
Dim Intentos As Integer
Dim PesoTara As Double
Dim Puerto As Integer
Dim Baudios As Long

Private Sub btnAceptarPeso_Click()

    Respuesta = "APROBO"
    Hide
    
End Sub

Private Sub btnCancelar_Click()
    
    TXTPESO.TEXT = ""
    Hide

End Sub

Private Sub btnRefresh_Click()

    lblLectura.Caption = "Leyenda Balanza ... ": lblLectura.Refresh
        
    tmrObtienePesada.Interval = 1000
    tmrObtienePesada.Enabled = True
    
End Sub

'
Private Sub Form_Load()

    Respuesta = "CANCELO"
    
    ' CALCULO TARA DEL ARTICULO
    PesoTara = 0
    TaraPesoUnMetro# = XVALO(CONTROL$("BOLRECEP", "TARAMET"))
    Puerto = XVALO(CONTROL$(IDENTER$, "PORTBAL"))
    Baudios = XVALO(CONTROL$(IDENTER$, "BAUDBAL"))
    PesoTara = TaraPesoUnMetro#
    
    ' APERTURA DE PUERTO COM
    
    On Error GoTo AperturaPuertoCOM
    
    'Indicamo en numero del puerto a usar
    MSComm1.CommPort = Puerto
    
    'Otros seteos basicos
    MSComm1.settings = CStr(Baudios) & ",N,8,1"
    MSComm1.InBufferSize = 32   ' Establece y devuelve el tamaño del búfer de recepción, en bytes
    MSComm1.OutBufferSize = 512 ' Establece y devuelve el tamaño, en bytes, del búfer de transmisión.
    MSComm1.RThreshold = 32     ' Cant de caracteres cuando se activa evento oncomm y asigna a CommEvent el valor comEvReceive (0 desactiva >=1 activa)
    MSComm1.SThreshold = 1      ' Idem anterior pero para transmision de caracteres. Asigna a CommEvent el valor comEvSend
    MSComm1.InputLen = 0        ' MSComm lee todo el contenido del búfer de recepción cuando se usa Input
    MSComm1.ParityReplace = "?"
    MSComm1.RTSEnable = True
    MSComm1.DTREnable = True
    MSComm1.NullDiscard = False 'Determina si se transfieren caracteres nulos desde el puerto al búfer de recepción.
    MSComm1.Handshaking = comNone
    'MSComm1.InputMode = comInputModeText 'Los datos se dan en modo texto
    
    ' Abrimos el puerto
    MSComm1.PortOpen = True
          
    On Error GoTo 0
    
    lblLectura.Caption = "Leyenda Balanza ... ": lblLectura.Refresh
    tmrObtienePesada.Interval = 1500
    tmrObtienePesada.Enabled = True
        
    Exit Sub
    
AperturaPuertoCOM:

    Call MENSERR(24, "Imposible Obtener Pesada. Error en la Apertura del Puerto (" & CStr(MSComm1.CommPort) & ")")
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    'Cancel = 1

End Sub

Private Sub MSComm1_OnComm()
    
    ' CAPTURA EVENTOS EMITIDOS POR EL MSCOMM Y ACTUA EN CONSECUENCIA

    ' declaramos una variable donde quedaran los datos recibidos
    Static strData As String
    
    With MSComm1
        Select Case .CommEvent
            Case comEventBreak
            MsgBox "Error", "1comEventBreak"
            Case comEventFrame
            MsgBox "Error", "1comEventFrame"
            Case comEventOverrun
            MsgBox "Error", "1comEventOverrun"
            Case comEventRxOver
            MsgBox "Error", "1comEventRxOver"
            Case comEventRxParity
            MsgBox "Error", "1comEventRxParity"
            Case comEventTxFull
            MsgBox "Error", "1comEventTxFull"
            Case comEventDCB
            MsgBox "Error", "1comEventDCB"
            Case comEvReceive
            
            ' se guardan los datos en la variable
            strData = strData & MSComm1.Input
            txtLecturaContinua = strData
            
            ' borramos la variable para los proximos datos
            strData = ""
        End Select
    End With
    
End Sub

Private Sub tmrObtienePesada_Timer()
    
    ' OBTIENE DE LOS DATOS RECIBIDOS LA PESADA
    tmrObtienePesada.Enabled = False
    TXTPESO.Enabled = False
    btnAceptarPeso.Enabled = False
    btnCancelar.Enabled = False
    btnRefresh.Enabled = False
    Intentos% = 0: INI% = 1
    
10  Intentos% = Intentos% + 1
    lblIntentos.Caption = CStr(Intentos%)
    
    LecturaBalanza$ = txtLecturaContinua.TEXT
    
    If Intentos% >= 20 Then
        lblLectura.Caption = "Lectura Fallida": lblLectura.Refresh
        GoTo 90
    End If
    
    If TRIM$(LecturaBalanza$) = "" Then
        GoTo 10
    End If

    CaracStatus$ = "D" ' STATUS: PESO BRUTO MAYOR A CERO
    CaracFinPalabra$ = Chr$(13)
    
    n0% = InStr(INI%, LecturaBalanza$, CaracStatus$)
    If n0% > 0 Then
        N1% = InStr(n0% + 1, LecturaBalanza$, CaracFinPalabra$)
        If N1% > 0 And N1% > n0% Then
            PESOLEIDO# = XVALO(Mid$(LecturaBalanza$, n0% + 1, N1% - n0% - 1)) / 10
            'PesoLeido# = PesoLeido# - PesoTara
            If PESOLEIDO# > 0.005 Then
                TXTPESO.TEXT = TRIM$(FORMATNUM$(PESOLEIDO#, "F10.1"))
            
                lblLectura.Caption = "Lectura OK": lblLectura.Refresh
            Else ' No obtuvo el peso entre los simbolos (¢)
                'ini% = n1% + 1
                GoTo 10
            End If
        Else ' No encontro el caracter 13 de fin de peso
            'ini% = n0% + 1
            GoTo 10
        End If
    Else ' No encontro el primer simbolo (¢) a partir del cual sigue el peso
        'ini% = Len(LecturaBalanza%) + 1
        GoTo 10
    End If
    
90  TXTPESO.Enabled = True
    btnAceptarPeso.Enabled = True
    btnCancelar.Enabled = True
    btnRefresh.Enabled = True
    
End Sub

Private Sub txtPeso_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 116 Then
        Call btnRefresh_Click
    End If

End Sub

Private Sub txtPeso_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        btnAceptarPeso.SetFocus
    End If
    
End Sub
