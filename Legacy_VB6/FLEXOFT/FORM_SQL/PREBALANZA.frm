VERSION 5.00
Begin VB.Form PREBALANZA 
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selección de Tara"
   ClientHeight    =   1800
   ClientLeft      =   150
   ClientTop       =   450
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   690
      TabIndex        =   2
      Top             =   240
      Width           =   175
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Balanza"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   75
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Lee el Peso de La Balanza"
      Top             =   1200
      Width           =   2190
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Kg. Tara"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3360
      TabIndex        =   7
      Top             =   840
      Width           =   975
   End
   Begin VB.Label lblkgtara 
      Alignment       =   1  'Right Justify
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
      Height          =   330
      Left            =   4080
      TabIndex        =   6
      Top             =   720
      Width           =   1185
   End
   Begin VB.Label LBLDESCTARA 
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
      Height          =   330
      Left            =   960
      TabIndex        =   5
      Top             =   240
      Width           =   4305
   End
   Begin VB.Label LBLTARA 
      Alignment       =   1  'Right Justify
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
      Height          =   330
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   555
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Selecionar Peso de Tara"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   3495
   End
   Begin VB.Label LblKgImpresos 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   2520
      TabIndex        =   1
      ToolTipText     =   "Kg. Con Tara Descontada"
      Top             =   1200
      Width           =   2745
   End
   Begin VB.Menu mnuT 
      Caption         =   "Transacciones"
   End
   Begin VB.Menu mnuUtil 
      Caption         =   "Utilidades"
   End
   Begin VB.Menu mnuConf 
      Caption         =   "Configuración"
      Begin VB.Menu mnuFormConsMat 
         Caption         =   "Formulario de Cosumo de Materiales"
      End
      Begin VB.Menu mnuTablaTaras 
         Caption         =   "Tabla de Taras"
      End
   End
End
Attribute VB_Name = "PREBALANZA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FORMULARIO_NAME$
Private Sub CMDCM_Click()


End Sub

Private Sub CMDDE_Click()

End Sub

Private Sub Command1_Click()
   X = LISTASELECGENERAL("TABLATARAS", , , 1)
   Me.LBLTARA = XVALO(RESP_ZELE_VECT(1))

'   Me.LBLDESCTARA = TRIM$(RESP_ZELE_VECT(2))
'   Me.lblkgtara = VALOBADA("DATASQL", "CAMPO3", "CONVERT(int, campo1) = '" & XVALO(RESP_ZELE_VECT(1)) & "' AND LISTA = 'TABLATARAS'", "NOMESS")

End Sub






Private Sub Command11_Click()
    Command11.Enabled = False
    
    If FORMULARIO_NAME$ = "" Then
       FORMULARIO_NAME$ = "TARASCRP"
    Else
       FORMULARIO_NAME$ = Left(UCase$(TRIM$(FORMULARIO_NAME$)), 8)
    End If
    
    Call GRACONTROL(IDENTER$, FORMULARIO_NAME$, Me.LBLTARA)
    Call GRACONTROL("*", FORMULARIO_NAME$, LBLTARA)
    

    Codigo$ = CODEXT$(CI1%)
    If XVALO(LBLTARA) < 1 Then
      Call COMUNI("Tara No Válido")
      GoTo 99
    End If
    
    TARASELEC% = XVALO(Me.LBLTARA)
    
    Dim CTP As New Conteplast
    PESOX# = CTP.PesoLeidoBalanza(0, TARASELEC%, OK%, 1)
    If PESOX# < 0 Then PESOX# = 0
    If OK% >= 0 Then
       LblKgImpresos = FORMATNUM$(PESOX#, "f12.2")
    Else
       LblKgImpresos = 0
    End If
99  Command11.Enabled = True
    Hide
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
    
    Me.LBLTARA = XVALO(CONTROL("*", "TARASCRP"))
    If CONTROL(IDENTER$, "TARASCRP") <> "" Then LBLTARA = XVALO(CONTROL(IDENTER$, "TARASCRP"))

End Sub


Private Sub LBLTARA_Change()
   Me.LBLDESCTARA = VALOBADA("DATASQL", "CAMPO2", "CONVERT(int, campo1) = '" & XVALO(LBLTARA) & "' AND LISTA = 'TABLATARAS'", "NOMESS")
   Me.lblkgtara = VALOBADA("DATASQL", "CAMPO3", "CONVERT(int, campo1) = '" & XVALO(LBLTARA) & "' AND LISTA = 'TABLATARAS'", "NOMESS")

End Sub





Private Sub mnuFormConsMat_Click()
    ConfigDepo.Show 1
End Sub





Private Sub mnuTablaTaras_Click()
    Dim TablaAuxiliar As New DatosAuxiliares
    Dim ListaCampos() As eDatosAuxiliaresCampos
    Dim CAMPO As eDatosAuxiliaresCampos
    
    ' TABLA DE TARAS
    ' **************
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Ord/I10"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "N"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    N = 1
    ReDim Preserve ListaCampos(N)
    Set ListaCampos(N - 1) = CAMPO
    
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Descripcion/A32"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "S"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    N = N + 1
    ReDim Preserve ListaCampos(N)
    Set ListaCampos(N - 1) = CAMPO
        
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Peso/F12.2"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "S"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    N = N + 1
    ReDim Preserve ListaCampos(N)
    Set ListaCampos(N - 1) = CAMPO
    
    Call TablaAuxiliar.CrearTabla("TABLATARAS", "Tabla de Taras", ListaCampos)

    
    IDLIST& = ObtenerIdTablaAuxiliar("TABLATARAS")
    Call ABM_TABLA(IDLIST&, "Tabla de Taras")

End Sub

