VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form RUTAWALD 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Conexión "
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   13650
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   13650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command9 
      Caption         =   "Consultar Sobre Remitos Generados en Conteplast"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   8760
      TabIndex        =   21
      Top             =   8520
      Width           =   3735
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      Left            =   120
      TabIndex        =   17
      Top             =   3680
      Width           =   12360
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   9000
         MaxLength       =   1
         TabIndex        =   26
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5640
         MaxLength       =   1
         TabIndex        =   24
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2640
         MaxLength       =   1
         TabIndex        =   22
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox Text3 
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
         Height          =   795
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Text            =   "RUTAWALD.frx":0000
         Top             =   480
         Width           =   12000
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Seleccionar"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   10560
         TabIndex        =   18
         Top             =   1320
         Width           =   1455
      End
      Begin MSComDlg.CommonDialog CommonDialog2 
         Left            =   1200
         Top             =   2160
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Prefijo Remtos Quilmes en Waldbott:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   6600
         TabIndex        =   27
         Top             =   1440
         Width           =   2475
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Prefijo Remtos Cambaceres en Waldbott:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3240
         TabIndex        =   25
         Top             =   1440
         Width           =   2475
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Prefijo Remtos Conteplast en Waldbott:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   120
         TabIndex        =   23
         Top             =   1395
         Width           =   2475
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "RUTA CONEXION A CARPETA DONDE ESTAN LOS AARCHIVOS GENERADOS EN CONTEPLAST PARA LEER Y LLEVAR A MOVISTOCTP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Width           =   11835
      End
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Consultar  Sobre Ruta donde se Leen los  Clientes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   4560
      TabIndex        =   16
      Top             =   8520
      Width           =   3735
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1875
      Left            =   120
      TabIndex        =   12
      Top             =   1800
      Width           =   12360
      Begin VB.CommandButton Command5 
         Caption         =   "Seleccionar"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   10560
         TabIndex        =   15
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox Text2 
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
         Height          =   915
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Text            =   "RUTAWALD.frx":0006
         Top             =   480
         Width           =   12000
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   1200
         Top             =   2160
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "RUTA CONEXION A WALDBOT (  ++++DONDE SE LEEN LOS CLIENTES +++++ )   W:\ENVAPLAST\SGIW11.MDB"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   11835
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Consultar  Sobre Ruta donde se Acceso a los remitos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   120
      TabIndex        =   11
      Top             =   8520
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Text            =   "RUTAWALD.frx":0008
      Top             =   7800
      Width           =   12360
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   2055
      Left            =   120
      TabIndex        =   9
      Top             =   5760
      Width           =   12360
      _ExtentX        =   21802
      _ExtentY        =   3625
      _Version        =   393216
      AllowUserResizing=   1
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Seleccionar"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   10800
      TabIndex        =   8
      Top             =   1440
      Width           =   1455
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   8775
      Left            =   12600
      ScaleHeight     =   8715
      ScaleWidth      =   1185
      TabIndex        =   1
      Top             =   -75
      Width           =   1245
      Begin VB.CommandButton command1 
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
         Height          =   700
         Left            =   150
         Picture         =   "RUTAWALD.frx":000A
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   285
         Width           =   800
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Guardar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   150
         Picture         =   "RUTAWALD.frx":0154
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1200
         Width           =   800
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Trasla."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   0
         Picture         =   "RUTAWALD.frx":045E
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Generar Diskette de Transferencia"
         Top             =   9840
         Visible         =   0   'False
         Width           =   800
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Imprimir"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   0
         Picture         =   "RUTAWALD.frx":0768
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Aprueba - Graba e Imprime Remito"
         Top             =   8880
         Width           =   800
      End
      Begin VB.Image Image2 
         Height          =   510
         Left            =   -405
         Picture         =   "RUTAWALD.frx":0DD2
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   8280
         Width           =   2010
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12360
      Begin MSComDlg.CommonDialog cmd 
         Left            =   1200
         Top             =   2160
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox Text6 
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
         Height          =   795
         Left            =   100
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Text            =   "RUTAWALD.frx":43BC
         Top             =   480
         Width           =   12000
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "RUTA CONEXION A WALDBOT (  ++++DONDE SE ESCRIBEN LOS REMITOS+++++ )  W:\ENVAPLASTFC\SGIW11.MDB"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   11715
      End
   End
End
Attribute VB_Name = "RUTAWALD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub


Private Sub Command2_Click()
    '
    Call GRACONTROL("CONWALDB", "RUTAWALD", Text6)
    Call GRACONTROL("CONWALDB", "RTLECTCL", Text2) 'LECUTRA DE CLIENTES
    Call GRACONTROL("TRCTPAWL", "DROPTRRM", Text3) 'RUTA CONEXION A CARPETA DE INTERCAMBIO DE REMITOS GENERADOS EN CONTEPLAST PARA PASAR A WALDBOTT
    Call GRACONTROL("TRCTPAWL", "PFIJOCTP", Text4) 'PREFIJO CONTEPLAST REMITOS
    Call GRACONTROL("TRCAPAWL", "PFIJOCAM", Text5) 'PREFIJO CAMBACERES REMITOS
    Call GRACONTROL("TRCAPAWL", "PFIJOQUI", Text7) 'PREFIJO QUILMES REMITOS
    Call CIERRARCH("*.*")
    Hide
    '
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   RUTAX$ = TRIM$(SELECCION_Archivo$("*.*", cmd))
   If RUTAX$ <> "" Then Text6 = SELECCION_Archivo$("*.*", cmd)
   Command3.Enabled = True
   Screen.MousePointer = 1
End Sub
'
Private Sub Command4_Click()
    Command4.Enabled = False
    On Error Resume Next
    Call abreBaDaWald
    If Err > 0 Then Call COMUNI("Imposible Conectar con Base de Datos"): GoTo 99
    On Error GoTo 0
    If TRIM$(Text1) = "" Then Call COMUNI("Falta Ingresar Consulta"): GoTo 99
    MSF.Rows = 1
    SQLX$ = Text1
    Dim rs As New ADODB.Recordset
    Set rs = BASEWALD.Execute(SQLX$)
    
    J% = 0
    With rs
    Do While J% < .Fields.Count
      On Error Resume Next
      MSF.TextMatrix(0, J%) = rs.Fields(J%).Name
      If Err Then Exit Do
      On Error GoTo 0
      J% = J% + 1
      MSF.Cols = J% + 1
'     .MoveNext
    Loop
   End With
   MSF.Cols = J% 'de este modo le resto 1 si no queda uno de mas sin nombre de campo
   
'    Set RS = Nothing
     i& = 0: jj% = 0
     SQLX$ = Text1
     '
     'CARGA LOS VALORES.
'     Set PARATEMP = BaDaWald.OpenRecordset(SQLX$, dbOpenSnapshot)
     With rs
       Do While Not .EOF
         i& = i& + 1: jj% = jj% + 1
         MSF.Rows = i& + 1
         For jj% = 0 To J% - 1
           If IsNull(rs.Fields(jj%).Value) Then GoTo 60
            MSF.TextMatrix(i&, jj%) = rs.Fields(jj%).Value
60       Next jj%
         .MoveNext
       Loop
     End With
     rs.Close
     Set rs = Nothing
    
   
99 Command4.Enabled = True
End Sub


Private Sub Command5_Click()
   Command5.Enabled = False
   RUTAX$ = TRIM$(SELECCION_Archivo$("*.*", cmd))
   If RUTAX$ <> "" Then Text2 = SELECCION_Archivo$("*.*", cmd)
   Command5.Enabled = True
   Screen.MousePointer = 1
End Sub

Private Sub Command7_Click()
    Command7.Enabled = False
    On Error Resume Next
    Call abreBaDaWaldClientes
    If Err > 0 Then Call COMUNI("Imposible Conectar con Base de Datos"): GoTo 99
    On Error GoTo 0
    If TRIM$(Text1) = "" Then Call COMUNI("Falta Ingresar Consulta"): GoTo 99
    MSF.Rows = 1
    SQLX$ = Text1
    Dim rs As New ADODB.Recordset
    Set rs = BASEWALDCLIENTE.Execute(SQLX$)
    
    J% = 0
    With rs
    Do While J% < .Fields.Count
      On Error Resume Next
      MSF.TextMatrix(0, J%) = rs.Fields(J%).Name
      If Err Then Exit Do
      On Error GoTo 0
      J% = J% + 1
      MSF.Cols = J% + 1
'     .MoveNext
    Loop
   End With
   MSF.Cols = J% 'de este modo le resto 1 si no queda uno de mas sin nombre de campo
   
'    Set RS = Nothing
     i& = 0: jj% = 0
     SQLX$ = Text1
     '
     'CARGA LOS VALORES.
'     Set PARATEMP = BaDaWald.OpenRecordset(SQLX$, dbOpenSnapshot)
     With rs
       Do While Not .EOF
         i& = i& + 1: jj% = jj% + 1
         MSF.Rows = i& + 1
         For jj% = 0 To J% - 1
           If IsNull(rs.Fields(jj%).Value) Then GoTo 60
            MSF.TextMatrix(i&, jj%) = rs.Fields(jj%).Value
60       Next jj%
         .MoveNext
       Loop
     End With
     rs.Close
     Set rs = Nothing
     BASEWALDCLIENTE.Close
99 Command7.Enabled = True

End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   RUTAX$ = TRIM$(SELECCION_Archivo$("*.*", cmd))
   If RUTAX$ <> "" Then Text3 = SELECCION_Archivo$("*.*", cmd)
   Command8.Enabled = True
   Screen.MousePointer = 1

End Sub

Private Sub Command9_Click()
    Command9.Enabled = False
'    On Error Resume Next
''    Call abreBaMDBINTER_REMI_ENVAP
'
'    If Err > 0 Then Call COMUNI("Imposible Conectar con Base de Datos"): GoTo 99
'    On Error GoTo 0
    If TRIM$(Text1) = "" Then Call COMUNI("Falta Ingresar Consulta"): GoTo 99
    MSF.Rows = 1
    SQLX$ = Text1
'    Dim RS As New ADODB.Recordset
    Set rs = READSET(SQLX$)
    
    J% = 0
    With rs
    Do While J% < .Fields.Count
      On Error Resume Next
      MSF.TextMatrix(0, J%) = rs.Fields(J%).Name
      If Err Then Exit Do
      On Error GoTo 0
      J% = J% + 1
      MSF.Cols = J% + 1
'     .MoveNext
    Loop
   End With
   MSF.Cols = J% 'de este modo le resto 1 si no queda uno de mas sin nombre de campo
   
'    Set RS = Nothing
     i& = 0: jj% = 0
     SQLX$ = Text1
     '
     'CARGA LOS VALORES.
'     Set PARATEMP = BaDaWald.OpenRecordset(SQLX$, dbOpenSnapshot)
     With rs
       Do While Not .EOF
         i& = i& + 1: jj% = jj% + 1
         MSF.Rows = i& + 1
         For jj% = 0 To J% - 1
           If IsNull(rs.Fields(jj%).Value) Then GoTo 60
            MSF.TextMatrix(i&, jj%) = rs.Fields(jj%).Value
60       Next jj%
         .MoveNext
       Loop
     End With
     rs.Close
     Set rs = Nothing
'     BASEINTERREMI_ENVAP.Close
99 Command9.Enabled = True



End Sub

Private Sub Form_Load()
    '
    Text6 = Control("CONWALDB", "RUTAWALD")
    Text2 = Control("CONWALDB", "RTLECTCL")
    Text3 = Control("TRCTPAWL", "DROPTRRM") 'RUTA CONEXION A CARPETA DE INTERCAMBIO DE REMITOS GENERADOS EN CONTEPLAST PARA PASAR A WALDBOTT
    Text4 = XVALO(Control("TRCTPAWL", "PFIJOCTP")) 'PREFIJO CONTEPLAST REMITOS
    Text5 = XVALO(Control("TRCAPAWL", "PFIJOCAM")) 'PREFIJO CAMBACERES REMITOS
    Text7 = XVALO(Control("TRCAPAWL", "PFIJOQUI")) 'PREFIJO CAMBACERES REMITOS
    
    '
End Sub











Private Sub Text6_DblClick()
  Text6 = ""
End Sub
