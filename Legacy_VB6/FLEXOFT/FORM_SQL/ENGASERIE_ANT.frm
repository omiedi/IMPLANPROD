VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form ENGASERIE_ANT 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asociación de Números de Serie a Ordenes de Fabricación  (Solo se Podrán Relacionar N.Serie que sean parte de su Estructura)"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   12345
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6300
   ScaleWidth      =   12345
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command7 
      Caption         =   "Guardar Ingresos de N° de Series"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   29
      Top             =   5280
      Width           =   3435
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Imprimir Etiquetas (Grde.)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9060
      TabIndex        =   27
      ToolTipText     =   "Imprime las Etiquetas de los N° de Series Seleccionados"
      Top             =   2100
      Width           =   2955
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00808080&
      Height          =   975
      Left            =   120
      TabIndex        =   22
      Top             =   6120
      Width           =   5950
      Begin VB.CommandButton Command5 
         Caption         =   "Impresión de Etiquetas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3120
         TabIndex        =   26
         Top             =   240
         Width           =   2595
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "QR- Orden de Fabricación"
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
         Left            =   120
         TabIndex        =   25
         Top             =   120
         Width           =   2715
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Etiq. Pequeaña de N° Serie"
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
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Width           =   2475
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00E0E0E0&
         Caption         =   " N° Orden Series Generadas"
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
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   2715
      End
   End
   Begin VB.CommandButton Command3 
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
      Height          =   375
      Left            =   9120
      TabIndex        =   18
      Top             =   5280
      Width           =   2955
   End
   Begin VB.CommandButton Command4 
      Caption         =   ">>>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6000
      TabIndex        =   13
      Top             =   3840
      Width           =   200
   End
   Begin VB.CommandButton Command2 
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
      Height          =   300
      Left            =   2760
      TabIndex        =   10
      Top             =   120
      Width           =   175
   End
   Begin VB.CommandButton Command1 
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
      Left            =   13800
      TabIndex        =   3
      Top             =   7320
      Visible         =   0   'False
      Width           =   1600
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7680
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2160
      OleObjectBlob   =   "ENGASERIE_ANT.frx":0000
      Top             =   4920
   End
   Begin MSFlexGridLib.MSFlexGrid MSFSERIEVINCULADOS 
      Height          =   1845
      Left            =   6240
      TabIndex        =   0
      Top             =   240
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   3254
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      FillStyle       =   1
      GridLines       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Roboto Mono"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSFSRIESDISPONIBLES 
      Height          =   4005
      Left            =   360
      TabIndex        =   12
      Top             =   1200
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   7064
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      FillStyle       =   1
      GridLines       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Roboto Mono"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSFSELECCIONADOS 
      Height          =   2685
      Left            =   6240
      TabIndex        =   16
      Top             =   2520
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   4736
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      FillStyle       =   1
      GridLines       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Roboto Mono"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label LBLCI1DODIORFA 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   11400
      TabIndex        =   28
      Top             =   0
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Label lblcodvinculado 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   10440
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   4560
      TabIndex        =   20
      Top             =   840
      Width           =   1545
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Código:"
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
      Index           =   4
      Left            =   3720
      TabIndex        =   19
      Top             =   885
      Width           =   765
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Series Seleccionadas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   3
      Left            =   6240
      TabIndex        =   17
      Top             =   2280
      Width           =   1965
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cant."
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
      Index           =   0
      Left            =   4560
      TabIndex        =   15
      Top             =   165
      Width           =   525
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   5160
      TabIndex        =   14
      Top             =   120
      Width           =   945
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Series Vinculados"
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
      Index           =   1
      Left            =   6270
      TabIndex        =   11
      Top             =   30
      Width           =   2205
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "N° de Serie Disponibles:"
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
      Index           =   5
      Left            =   360
      TabIndex        =   9
      Top             =   870
      Width           =   2325
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   360
      TabIndex        =   8
      Top             =   480
      Width           =   5745
   End
   Begin VB.Label Label44 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2985
      TabIndex        =   7
      Top             =   120
      Width           =   1545
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nro. O/F"
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
      Index           =   2
      Left            =   0
      TabIndex        =   6
      Top             =   165
      Width           =   1125
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1185
      TabIndex        =   5
      Top             =   120
      Width           =   1545
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      BeginProperty Font 
         Name            =   "Roboto Mono"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   0
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Label5"
      Height          =   240
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      BeginProperty Font 
         Name            =   "Roboto Mono"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Menu MNUARCHIVO 
      Caption         =   "Archivo"
      Visible         =   0   'False
      Begin VB.Menu mnucopiar 
         Caption         =   "Copiar Celda Seleccionada"
         Visible         =   0   'False
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuitar 
         Caption         =   "Quitar"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnxarchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mnuReGenerarReserva 
         Caption         =   "Re-Generar Reservas de n.Serie Activo"
      End
      Begin VB.Menu mnuImpEtiqGde 
         Caption         =   "Imprimir Etiqueta de N.Serie (Grande)"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu MnuEdición 
      Caption         =   "Edición"
      Begin VB.Menu mnuMenu3 
         Caption         =   "Revertir Selección"
      End
      Begin VB.Menu s2 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuConfig 
      Caption         =   "Configuración"
      Begin VB.Menu mnuControlesxequipos 
         Caption         =   "Controles x Equipos"
      End
      Begin VB.Menu mnuConfigEtiquetas 
         Caption         =   "Configuraciones de Formularios de Etiquetas"
      End
   End
End
Attribute VB_Name = "ENGASERIE_ANT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim GLOBALFRM_strContenidoCelda$
Dim GLOBALFRM_FILA&
Dim MSFGLOBALFRM As MSFlexGrid

Dim LecturaDatosOriginales As Byte
Public MENU1%
Public MENU2%
Public MENU3%
Public NOPINTAR%
Public SeleccionaFilaDobleClick%
Public ESCODIGOBASE%

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Sub ACTUALIZO_TODO_TRAZAXSERIE(NSERIEORIGINAL$, NUMSERINUEVO$, ID_ACTUALIZADO_TANUMSE, CI1%)
    'EJEMPLO: SI EL USUARIO ASOCIA EL N.SERIE ORIGINAL ABC005 AL SIGUIENTE DE LA ESTRUCTURA Y AHORA SE GENERA EL NUEVO
    'N.SERIE ABC001 CAMBIA EL ABC005 POR EL ABC001 Y SI EL ABC001 EXISTIA ANTERIORMENTE LE APLICA LO CAMBIA POR ABC005

    CONDI$ = "NumeroSerie= '" & NUMSERINUEVO$ & "'"
    SQLX$ = "SELECT TRAZAXSERIE_ID FROM TRAZAXSERIE WITH(NOLOCK) WHERE " & CONDI$
    Set RS = READSET(SQLX$)
    i& = 0
    With RS
      Do While Not .EOF
          ID& = XVALO(!TRAZAXSERIE_ID)
          i& = i& + 1
          If i& = 1 Then
            CONDI$ = " (TRAZAXSERIE_ID = " & ID&
          Else
            CONDI$ = CONDI$ + " or TRAZAXSERIE_ID = " & ID&
          End If
        .MoveNext
      Loop
    End With
    CONDI$ = CONDI$ + ")"
    
    If i& < 0 Then Exit Sub
    
    CONDI2$ = "NumeroSerie= '" & NSERIEORIGINAL$ & "'"
    Call ACTUREGISTRO("TRAZAXSERIE", "NumeroSerie", CONDI2$, NUMSERINUEVO$, REGAF&, "NOMESS")
    
'    Call ACTUREGISTRO("TRAZAXSERIE", "NumeroSerie", CONDI$, NSERIEORIGINAL$, REGAF&, "NOMESS")
    
    For J& = 1 To MSFSRIESDISPONIBLES.Rows - 1
      If TRIM$(MSFSRIESDISPONIBLES.TextMatrix(J&, 3)) = NUMSERINUEVO$ Then
       Call ACTUREGISTRO("TRAZAXSERIE", "NumeroSerie", CONDI$, NSERIEORIGINAL$, REGAF&, "NOMESS")
       MSFSRIESDISPONIBLES.TextMatrix(J&, 3) = NSERIEORIGINAL$
       If ID_ACTUALIZADO_TANUMSE > 0 Then
            CONDI$ = "COD_INTE= " & CI1% & " AND NUMEROSERIE = '" & NUMSERINUEVO$ & "' AND TANUMSE_ID <> " & ID_ACTUALIZADO_TANUMSE
            Call ACTUREGISTRO("TANUMSE", "NUMEROSERIE", CONDI$, NSERIEORIGINAL$, REGAF&, "NOMESS")
       End If
       Exit For
      End If
    Next J&

End Sub

Sub CargarNSerie_aCodigoBase(NORFA$, CI1%, MSF As MSFlexGrid, FEALTA%)
    'TRAER LOS QUE EXISTEN EN LA OF Y INNER CON TRAZAXSERIE
    Dim NSERIES$(): Dim ID&()
    CANP% = CantiProyectadaOF(NORFA$)
    fech$ = FECHATEX$(FEALTA%)
    J% = 0
    SQLX$ = "SELECT NUMEROSERIE,TRAZAXSERIE_ID "
    SQLX$ = SQLX$ + " FROM TRAZAXSERIE WITH(NOLOCK)  "
    SQLX$ = SQLX$ + " WHERE COD_INTE = " & CI1%
    SQLX$ = SQLX$ + " AND IDSUBOR = '" & NORFA$ & "'"
    Set RS = READSET(SQLX$)
    With RS
        Do While Not .EOF
           J% = J% + 1
           ReDim Preserve NSERIES$(J%), ID&(J%)
           NSERIES$(J%) = SAFETEXT$(!NUMEROSERIE)
           ID&(J%) = XVALO(!TRAZAXSERIE_ID)
           .MoveNext
        Loop
    End With
    '
       For i% = 1 To CANP%
          MSF.Rows = i% + 1
          
          MSF.TextMatrix(i%, 1) = fech$
          MSF.TextMatrix(i%, 2) = NORFA$
          If i% <= J% Then
            MSF.TextMatrix(i%, 3) = NSERIES$(i%) 'SI TIENE N.SERIE LO VA PONIENDO EN LOS PRIMEROS
            MSF.TextMatrix(i%, 4) = ID&(i%)
          Else
            MSF.TextMatrix(i%, 3) = "" ' SI NO GENERA  VACIO
            MSF.TextMatrix(i%, 4) = ""
          End If
       Next i%
End Sub

Sub LIMPIAR()
   Me.MSFSERIEVINCULADOS.Rows = 1
   Me.MSFSRIESDISPONIBLES.Rows = 1
   Me.MSFSELECCIONADOS.Rows = 1
   Label3.Caption = ""
   Label4.Caption = ""
   Label6.Caption = ""
   Label44.Caption = ""
   Me.LBLCI1DODIORFA = ""
   
End Sub

Private Sub Command2_Click()

   Command2.Enabled = False
   On Error Resume Next
   RESP_ZELE_VECT(1) = ""
   On Error GoTo 0
   CAMPOS$ = "IDSUBOR/A12;REFEREN/A48;C.Proy./F8.1"
   Call CARGA_ENCABEZADO(FRMZELECHO.msf2, CAMPOS$, FRMZELECHO)
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, CAMPOS$, FRMZELECHO)
   FRMZELECHO.Caption = ""
 
   Dim obj As New RECORXET
   Set RS1 = obj.RS_OF_ABIERTAS(1)
   Call Carga_MSF_Recordset(RS1, CAMPOS$, FRMZELECHO.msf2)
   Set obj = Nothing
   RS1.Close
   Set RS1 = Nothing

   FRMZELECHO.Width = 9500
   FRMZELECHO.Show 1
   If TRIM$(RESP_ZELE_VECT(1)) <> "" Then
      Label13 = TRIM$(RESP_ZELE_VECT(1))
      Label4 = TRIM$(RESP_ZELE_VECT(3))
   End If
    '
99  Command2.Enabled = True

End Sub



Private Sub Command3_Click()
       Command3.Enabled = False
       NORFA$ = Label13
       CI1% = LBLCI1DODIORFA
       If MSFSELECCIONADOS.Rows <= 1 Then
          Call COMUNI("No se Detectaron N.Series Seleccionados")
          GoTo 99
       End If
       '
       If (MSFSELECCIONADOS.Rows - 1 + Me.MSFSERIEVINCULADOS.Rows - 1) > XVALO(Label4) Then
          Call COMUNI("La Suma de Números de Serie Seleccionados\Más la Cantidad Yá Asignada\Supera la Cantidad De Equipos Para La Orden Activa")
          GoTo 99
       End If
       
       
'       ESEXCEPCION_SERIE% = ES_EXCEPCION_GENERAR_SERIE%(Label6) ' SI ESTA EN LA LISTA (DATASQL) DEVUELVE EL CODIGO INTERNO
       On Error GoTo ERROR_GUARDAR
       FLEXCONN.BeginTrans

        
       For i& = 1 To Me.MSFSELECCIONADOS.Rows - 1
          NS$ = MSFSELECCIONADOS.TextMatrix(i&, 3)
          If i& = 1 Then
            CONDI$ = " AND (NUMEROSERIE = '" & NS$ & "'"
          Else
            CONDI$ = CONDI$ + " OR NUMEROSERIE = '" & NS$ & "'"
          End If
       Next i&
       CONDI$ = CONDI$ + ")"
       '
       SQLX$ = "SELECT OF_ARRASTRA FROM TRAZAXSERIE  "
       SQLX$ = SQLX$ + " Where COD_INTE = " & XVALO(lblcodvinculado) & " AND (OF_ARRASTRA IS NULL OR OF_ARRASTRA = '')"
       SQLX$ = SQLX$ + CONDI$
       Dim RS As ADODB.Recordset
       Set RS = New ADODB.Recordset
25     On Error Resume Next
       RS.Open (SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    
       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
          On Error GoTo 0
          Call CapturaErrorOpen
          GoTo 25
       End If
       On Error GoTo 0
       '
       U& = 0
       With RS
          Do While Not .EOF
             !OF_ARRASTRA = Label13
             .Update
             .MoveNext
          Loop
       End With
       '
60     RS.Close
       Set RS = Nothing
       Screen.MousePointer = 1
       '
       For i& = 1 To Me.MSFSELECCIONADOS.Rows - 1
          NS$ = MSFSELECCIONADOS.TextMatrix(i&, 3)
          Call GENERATRAZAXSERIEM1(NORFA$)
'          Call AltaReservaEnTrazaxSerie_2(Label13, CODINT%(Label44), NS$, 1)
'
          Call AltaEnTrazaXSerieSistemaAnterior(NORFA$, CI1%, NS$) 'ALTA DE ITEM EN TRAZAXSERIE
          Call AltaReservaEnTrazaxSerie_SistemaAnterior(NORFA$, CI1%, NS$) 'ALTA DE RESERVAS EN TRAZAXSERIE

       Next i&
       
       
ERROR_GUARDAR:
    If Err.Number <> 0 Then
       FLEXCONN.RollbackTrans
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description)
       On Error GoTo 0
    Else
      FLEXCONN.CommitTrans
      Call COMUNI("Proceso Completo")

      On Error GoTo 0
      
    End If
    Call Label13_Change
       
99  Command3.Enabled = True

'       VER ACA CUANDO SE GRABA DE GENERAR LA GRABACION PARA LOS NUMEROS DE SERIE
'       VER SI SE PUEDE CONTROLAR QUE ESTO SE APLIQUE SOLO A LOS QUE SE HAN ASIGNADO SU NUMEROS DE SERIE.
End Sub


Private Sub Command4_Click()
  With MSFSRIESDISPONIBLES
   For i& = 1 To .Rows - 1
      MARCA$ = .TextMatrix(i&, 0)
      If MARCA$ = "*" Then
        Call AGREGAR_FILA_OTRO_MSF(MSFSRIESDISPONIBLES, Me.MSFSELECCIONADOS, i&)
      End If
   Next i&
  End With
  Call ELIMINADUPLICADOS(MSFSELECCIONADOS, 3)
End Sub

Private Sub Command5_Click()
   
    Call IMPRIMIR_ETIQUETAS_EYM(MSFSERIEVINCULADOS, Label13, Check1.Value, Check2.Value, Check3.Value)
End Sub

Private Sub Command6_Click()
        Command6.Enabled = False
        PIMPRE$ = CONTROL$(IDENTER$, "PORTSTIK")
        If PIMPRE$ = "" Then PIMPRE$ = CONTROL$("*", "PORTSTIK")
        NORFA$ = Label13
        
' Call SETPRINPORT(PIMPRE$)
'
'        Call PRINTDOC("@" & "PRUE-EYM")
'        Call SETPRINPORT("RESTORE")
        With MSFSERIEVINCULADOS
         For i& = 1 To .Rows - 1
            MARCA$ = .TextMatrix(i&, 0)
            If MARCA$ = "*" Then
                NSERIE$ = .TextMatrix(i&, 3)
                FE% = FECHANUM(.TextMatrix(i&, 1))
                TEXTOQR$ = Label44 & " " & NSERIE$
                FILTX$ = NSERIE$ & ";" & Right(FETEXCOR1$(FE%), 4) & ";" & NORFA$
                Call IMPRIMIR_ETIQUETAQR(CONTROL("", "STIKAPRO"), FILTX$, TEXTOQR$, "IMAGNSERIE_QR.JPG", 1, "QR_NSERIE", PIMPRE$, ANCHOX%, ALTOX%)
            End If
         Next i&
        End With
99      Command6.Enabled = True

End Sub

Private Sub Command7_Click()
       Command7.Enabled = False
       Command3.Enabled = False
       NORFA$ = Label13
       CI1% = LBLCI1DODIORFA
       ''''''''''''''''''''''
       If MSFSRIESDISPONIBLES.Rows <= 1 Then
          Call COMUNI("Sin Datos")
          GoTo 99
       End If
       
       If ES_EXCEPCION_GENERAR_SERIE%(Label44) > 0 Then
          EXIX& = EXISTE_NSERIE_EN_TANUMSE%(Me.MSFSRIESDISPONIBLES, 3, 4)
          If EXIX& > 0 Then 'VALIDO QUE NO EXISTA EN LA TABLA TANUMSE
             Call COMUNI("Número de Serie: " & MSFSRIESDISPONIBLES.TextMatrix(EXIX&, 3))
             Call SELECCIONARFILA(MSFSRIESDISPONIBLES, EXIX&)
             GoTo 99
          Else
            For i& = 1 To MSFSRIESDISPONIBLES.Rows - 1
                ID = XVALO(MSFSRIESDISPONIBLES.TextMatrix(i&, 4))
                If ID > 0 Then GoTo 30

                NS$ = TRIM$(MSFSRIESDISPONIBLES.TextMatrix(i&, 3))
                If NS$ <> "" Then
                   Call ALTANSERIE(NORFA$, CI1%, Label3, NS$) 'ALTA EN TANUMSE
                   Call AltaEnTrazaXSerieSistemaAnterior(NORFA$, CI1%, NS$) 'ALTA DE ITEM EN TRAZAXSERIE
                   Call AltaReservaEnTrazaxSerie_SistemaAnterior(NORFA$, CI1%, NS$) 'ALTA DE RESERVAS EN TRAZAXSERIE
                End If
30           Next i&
          End If
       Else
          Call COMUNI("Solo Se Podrá Grabar Ingesos de N° de Series para El Nivel Inicial Más Bajo")
          GoTo 99
       End If
       
       Label13 = ""
       Label13 = NORFA$
99     Command7.Enabled = True
End Sub

Private Sub Form_Activate()
MENU1% = 0
MENU2% = 0
MENU3% = 0
SeleccionaFilaDobleClick% = 0
End Sub

Private Sub Form_Load()
'     CAMPOS$ = "O.F./A20; N.Serie/A20;ID/A0"

     CAMPOS$ = "Alta/D8;O.F./A20; N.Serie/A20;ID/A0"
     Call CARGA_ENCABEZADO(MSFSRIESDISPONIBLES, CAMPOS$, Me)

     
     Call CARGA_ENCABEZADO(Me.MSFSELECCIONADOS, CAMPOS$, Me)
     Call CARGA_ENCABEZADO(Me.MSFSERIEVINCULADOS, CAMPOS$, Me)
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If Cancel = 0 Then
     Unload Me
   End If
End Sub
Function REEMPLAZO_NS_TANUMSE(NSORIGINAL$, NSNUEVO$, MESANO$, CI1%, CODI$, DESCRIX$, NORFA$)
    ORDEX% = XVALO(Right$(NSNUEVO$, 3))
    TXTX$ = "Alta N/Serie a Partir de: " & NORFA$
    HOII% = HOY(HOS$)
    fech$ = HOS$
    Call FRATEXTO(TXTX$, 56) ' FRACCIONO EL TEXTO
    '
    For KKI% = 1 To CARETEX% ' RECORRO DESDE 1 A LA CANTIDAD DE FRACCIONES
       TTYY$ = Space$(96)
       If KKI% = 1 Then
          Call REPLA(TTYY$, fech$, 1, 8) ' SI ES LA PRIMERA VEZ GRABA LA FECHA Y USUARIO
          Call REPLA(TTYY$, USERTER$, 69, 24)
       End If
       Call REPLA(TTYY$, RETEX$(KKI%), 11, 56) ' SIEMPRE QUE PASA GUARDA EL CENTRO DEL TEXTO
       ANOTANTE$ = ANOTANTE$ + TTYY$ + vbCrLf
    Next KKI%
    ANOTANTE$ = ANOTANTE$ & String$(88, "=") & vbCrLf
    
    'HACE EL REEMPLAZO CON EL SP EN LA BASE DE DATOS Y DEVUELVE EL ID DONDE MODIFICO
    Dim OBJCMD As ADODB.Command 'El objeto
    Set OBJCMD = New ADODB.Command
    OBJCMD.CommandType = adCmdStoredProc
    OBJCMD.CommandText = "UPDATETANUMSE"
    Set OBJCMD.ActiveConnection = FLEXCONN
    Dim valor_NuevoID As Long

    'Los parametros
    OBJCMD.Parameters(1) = CodiEmpr%
    OBJCMD.Parameters(2) = NSORIGINAL$
    OBJCMD.Parameters(3) = NSNUEVO$
    OBJCMD.Parameters(4) = MESANO$
    OBJCMD.Parameters(5) = ORDEX%
    OBJCMD.Parameters(6) = CI1%
    OBJCMD.Parameters(7) = CODI$
    OBJCMD.Parameters(8) = DESCRIX$
    OBJCMD.Parameters(9) = NORFA$
    OBJCMD.Parameters(10) = FETEXCOR1$(HOY(HO$))
    OBJCMD.Parameters(11) = ANOTANTE$
    '
    'Mando a ejecutar el sp
    OBJCMD.Execute
    '
    'Capturo el valor de retorno
    valor_NuevoID = XVALO(OBJCMD.Parameters(0).Value)
    REEMPLAZO_NS_TANUMSE = valor_NuevoID
    


End Function



Private Sub GRID_Click()
'    Static MODO  As Boolean
'    i& = GRID.MouseRow
'    J& = GRID.MouseCol
'
'    If (GRID.MouseRow = 0) Then
'
'        ' Ordena en forma ascendente
'        If MODO Then
'            GRID.COL = GRID.MouseCol
'            GRID.Sort = 2
'            MODO = False
'        ' Ordena en forma descendente
'        Else
'            GRID.COL = GRID.MouseCol
'            GRID.Sort = 1
'            MODO = True
'        End If
'    ElseIf NOPINTAR% > 0 Then Exit Sub
'    End If
'    If J& = 0 Then
'       Call COLOREAR_GRILLA_SOLAFILA(GRID, vbYellow, GRID.Row, 0, 0, 0)
'    End If
End Sub



Private Sub Label13_Change()
   Call LIMPIAR
   NORFA$ = Label13
   
   CI1% = XVALO(VALOBADA("ORDEFABR", "Cod_Inte", "IdSubOr = '" & NORFA$ & "'"))
   FEALTA% = CVINT(VALOBADA("ORDEFABR", "FechGen", "IdSubOr = '" & NORFA$ & "'"))
   fech$ = FECHATEX$(FEALTA%)
   LBLCI1DODIORFA = CI1%
   Label4 = CantiProyectadaOF(NORFA$)
   If CI1% > 0 Then
     Label44 = CODEXT(CI1%)
     Label3 = DESCRIT$(CI1%)
     

     'BUSCO EL CODIGO DE LA ESTRUCUTRA QUE SEA TRAZABLE SERIE  -1 DEL CODIGO DE LA ORDEN ELEGIDA
     '============================
     SQLX$ = "SELECT CodXElem,CodiElem "
     SQLX$ = SQLX$ + " FROM FORMULAS F WITH(NOLOCK) INNER JOIN MASTER M WITH(NOLOCK)"
     SQLX$ = SQLX$ + " ON M.CODINT = CodiElem "
     SQLX$ = SQLX$ + " WHERE (M.TRAZABLE = 2 OR M.TRAZABLE = -1) AND CodIConj= " & CI1%
     Set RS = READSET(SQLX$)
     KODIELEM = XVALO(RS!CODIELEM)
     KODIXELEM = SAFETEXT$(RS!CODXELEM)
     RS.Close: Set RS = Nothing
     '============================
     '
     ESCODIGOBASE% = 0
     lblcodvinculado = KODIELEM
     Command4.Enabled = True: Command3.Enabled = True: Command7.Enabled = True
     If KODIELEM < 1 Then
        If ESTRAZABLE%(CI1%, 1) = 2 Then
          Command4.Enabled = False: Command3.Enabled = False:
          Call CargarNSerie_aCodigoBase(NORFA$, CI1%, Me.MSFSRIESDISPONIBLES, FEALTA%)
          ESCODIGOBASE% = 1
          Exit Sub
        Else
        End If
     Else
        Command7.Enabled = False

     End If
     
'     'BUSCO LOS NUMEROS DE SERIE DEL CODIGO QUE ENCONTRE EN LA CONSULTA INMEDIATAMENTE ARRIBA
'     'PRESENTO N.SERIE DISPONIBLES
'     i& = 0
'     Label6 = KODIXELEM
'     'SI ES EXCEPCION DE GENERAR NUMERO DE SERIE OCULTO LA COLUMNA POR QUE EN REALIDAD LO GENERA
'     ESEXCEPCION_SERIE% = ES_EXCEPCION_GENERAR_SERIE%(Label6)
'     If ESEXCEPCION_SERIE% > 0 Then
'        CAMPOS$ = "Alta/D8;O.F./A20; N.Serie/A0;ID/A0;Identificador/A16"
'        Call CARGA_ENCABEZADO(MSFSRIESDISPONIBLES, CAMPOS$, Me)
'        Call CARGA_ENCABEZADO(Me.MSFSELECCIONADOS, CAMPOS$, Me)
'     End If
     
     Label6 = KODIXELEM
     lblcodvinculado = KODIELEM
     Label4 = CantiProyectadaOF(NORFA$)
     
'     CAMPOS$ = "Alta/D8;O.F./A20; N.Serie/A20;ID/A0"
'     Call CARGA_ENCABEZADO(MSFSRIESDISPONIBLES, CAMPOS$, Me)
     SQLX$ = "SELECT FECHALTA,IDSUBOR,NUMEROSERIE,TRAZAXSERIE_ID FROM TRAZAXSERIE WITH(NOLOCK) "
     SQLX$ = SQLX$ + " Where COD_INTE = " & KODIELEM & " AND (OF_ARRASTRA IS NULL OR OF_ARRASTRA = '') ORDER BY IDSUBOR "
     Set RS = READSET(SQLX$)
     With RS
        Do While Not .EOF
           i& = i& + 1
           J& = J& + 1
           Me.MSFSRIESDISPONIBLES.Rows = i& + 1
           Me.MSFSRIESDISPONIBLES.TextMatrix(i&, 1) = SAFETEXT$(!FechAlta)
           Me.MSFSRIESDISPONIBLES.TextMatrix(i&, 2) = SAFETEXT$(!IDSUBOR)
           Me.MSFSRIESDISPONIBLES.TextMatrix(i&, 3) = SAFETEXT$(!NUMEROSERIE)
           Me.MSFSRIESDISPONIBLES.TextMatrix(i&, 4) = XVALO(!TRAZAXSERIE_ID)
           If ESEXCEPCION_SERIE% > 0 Then
              Me.MSFSRIESDISPONIBLES.TextMatrix(i&, 5) = Left$(SAFETEXT$(!IDSUBOR), 10) & J&
           End If
           .MoveNext
           If ESEXCEPCION_SERIE% > 0 Then
             If SAFETEXT$(!IDSUBOR) <> MSFSRIESDISPONIBLES.TextMatrix(i&, 2) Then J& = 0
           End If
        Loop
     End With
     RS.Close: Set RS = Nothing
     '
'    'CARGA DE VINCULADOS
'    '============================
     i& = 0
     CI1% = CODINT%(Label44)
     SQLX$ = "SELECT FECHALTA,IDSUBOR,NUMEROSERIE,TRAZAXSERIE_ID FROM TRAZAXSERIE WITH(NOLOCK) "
     SQLX$ = SQLX$ + " Where OF_ARRASTRA = '" & NORFA$ & "'"
     Set RS = READSET(SQLX$)
     With RS
        Do While Not .EOF
           i& = i& + 1
           Me.MSFSERIEVINCULADOS.Rows = i& + 1
           Me.MSFSERIEVINCULADOS.TextMatrix(i&, 1) = SAFETEXT$(!FechAlta)
           Me.MSFSERIEVINCULADOS.TextMatrix(i&, 2) = SAFETEXT$(!IDSUBOR)
           Me.MSFSERIEVINCULADOS.TextMatrix(i&, 3) = SAFETEXT$(!NUMEROSERIE)
           Me.MSFSERIEVINCULADOS.TextMatrix(i&, 4) = XVALO(!TRAZAXSERIE_ID)
           .MoveNext
        Loop
     End With
     RS.Close: Set RS = Nothing
  

   End If

End Sub


Private Sub mnuMenu1_Click()
  MENU1% = 1
  Hide
End Sub
Private Sub mnuMenu2_Click()
  MENU2% = 2
  Hide
End Sub

Private Sub mnuEngancharNSerieAOF_Click()
  
   ENGANNUMSERIE.Show 1
End Sub

 Sub INVICTRL()
   mnucopiar.Visible = Not (mnucopiar.Visible)
   mnuQuitar.Visible = Not (mnuQuitar.Visible)
End Sub
Sub MENU_CELDA(MSF As MSFlexGrid, REG&, COL%)
        If MSF.Rows <= 1 Then Exit Sub
        Call COLOREAR_GRILLA_SOLACELDA(MSF, vbYellow, REG&, Int(COL%))   ' LA PINTA DE AMARILLA
        Call INVICTRL
        GLOBALFRM_strContenidoCelda$ = TRIM$(MSF.TextMatrix(REG&, COL%))  'CELDA SELECCIONADA
        GLOBALFRM_FILA& = REG&
        PopupMenu MNUARCHIVO
        Call INVICTRL
        On Error Resume Next
        Call COLOREAR_GRILLA_SOLACELDA(MSF, vbWhite, REG&, COL%)  ' LA PINTA DE AMARILLA
        On Error GoTo 0
End Sub

Private Sub mnuConfigEtiquetas_Click()
   If DERACCE("CFGNUSE", "Configuracion de Formulario - N/Serie") > 1 Then
     CONETIEQEYM.Show 1
   End If

End Sub

Private Sub mnuControlesxequipos_Click()
    Unload VENTABNEW
    VENTANA.Inicializar = 0
    VENTANA.CampEditables = "2;4;6"
    VENTANA.AgregaRegistro = 1
    VENTANA.UtilizaFormula = 1
    VENTANA.HabilitaBorrar = 1
    VENTANA.HabilitaInsertar = 1
    VENTANA.NoMostrarBotonTilde = 0
    VENTANA.TITULO = "Selección de Equipos Que Llevan Control y Excepción (filas)"
    CAMPOS$ = "CI/A0;Código/A16;Descripción/A32;C.Control/A32;Referencia/A12;Excepciones (Filas)/A42"
    VENTANA.Campox = CAMPOS$
    VENTANA.ColEcoCampo(2) = "CODIGO;DESCRIPCION;MASTER"
    VENTANA.ColEcoCampo(4) = "CAMPO1;CAMPO2;DATASQL;LISTA  = 'CONFICONTROL'"
    VENTANA.Inicializar = 1
    '****************************************************************
    Dim obj As New RECORXET
    Set RS1 = obj.RS_EXCEPCIONES
    Call Carga_MSF_Recordset(RS1, CAMPOS$, VENTABNEW.MSF)
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    
30  VENTABNEW.Show 1
    Dim MSF As MSFlexGrid
    Set MSF = VENTABNEW.MSF
    '
    If VENTABNEW.APROBADO% > 0 Then
       On Error GoTo ERROR_GUARDAR
       FLEXCONN.BeginTrans
       Call BORRAREGISTROS("CONTROEYM", "COD_EXTE <> '' OR COD_EXTE = ''", REGAF&, "NOMESS")
       
       Dim RS As ADODB.Recordset
       SQLX$ = "SELECT * FROM CONTROEYM WHERE COD_EXTE <> ''"
       Set RS = New ADODB.Recordset
       RS.Open (SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText

       For i& = 1 To MSF.Rows - 1
          CI1% = CODINT%(MSF.TextMatrix(i&, 2))
          If CI1% > 0 Then
            COD$ = MSF.TextMatrix(i&, 2)
            Descrip$ = MSF.TextMatrix(i&, 3)
            CONT1$ = MSF.TextMatrix(i&, 4)
            CONT2$ = MSF.TextMatrix(i&, 5)
            EXCEPCIONES$ = MSF.TextMatrix(i&, 6)
            RS.AddNew
            RS!COD_INTE = CI1%
            RS!COD_EXTE = COD$
            RS!Descripcion = Descrip$
            RS!C_FRMCONTROL = CONT1$
            RS!REF_CONTROL = CONT2$
            RS!EXCEPCIONES = EXCEPCIONES$
            RS.Update
          End If
       Next i&
       '
ERROR_GUARDAR:
    If Err.Number <> 0 Then
       FLEXCONN.RollbackTrans
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description)
       RS.Close: Set RS = Nothing
       On Error GoTo 0
    Else
      FLEXCONN.CommitTrans
      RS.Close: Set RS = Nothing
    End If

    End If
End Sub

Private Sub mnuImpEtiqGde_Click()
   On Error Resume Next
   RESP_ZELE_VECT(1) = ""
   On Error GoTo 0
   CAMPOS$ = "CI/A0;Código/A16;Descripción/A28;O.Fabricación/A16;N.Serie/A12;F.Alta;Id/A0"
   Call CARGA_ENCABEZADO(FRMZELECHO.msf2, CAMPOS$, FRMZELECHO)
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, CAMPOS$, FRMZELECHO)
   FRMZELECHO.Caption = ""

   Dim obj As New RECORXET
   Set RS1 = obj.RS_SERIEEYM()
   Call Carga_MSF_Recordset(RS1, CAMPOS$, FRMZELECHO.msf2)
   Set obj = Nothing
   RS1.Close
   Set RS1 = Nothing

   FRMZELECHO.Width = 12500
   FRMZELECHO.Show 1
   If TRIM$(RESP_ZELE_VECT(4)) <> "" Then
   
'   Dim NSERIEX$(): Dim ORDEX$()
   For U& = 1 To Cantixxx&
           FILTX$ = NUMSERIDES$ & ";" & Right(FETEXCOR1$(FE%), 4)
           NUMSERIDES$ = RET_TOTAL_CON_CODI$ & CAHASTA$
           PIMPRE$ = CONTROL$(IDENTER$, "PORTSTIK")
           If PIMPRE$ = "" Then PIMPRE$ = CONTROL$("*", "PORTSTIK")
           Call IMPRIMIR_ETIQUETAQR(CONTROL("", "STIKAPRO"), FILTX$, NUMSERIDES$, "IMAGNSERIE_QR.JPG", 1, "QR_NSERIE", PIMPRE$, ANCHOX%, ALTOX%)
           'ACA SOLO SE IMPRIMEN LAS ETIQUETAS PEQUEÑAS
'           If Check3.Value = 1 Then
'               ReDim Preserve NSERIEX$(U&) ' NUMERO DE SERIE
'               NSERIEX$(U&) = NUMSERIDES$
'           End If
'           If Check2.Value = 1 Then
'               ReDim Preserve ORDEX$(U&)
'               ORDEX$(U&) = Right(NUMSERIDES$, 3) ' NUMERO DE ORDEN DE SERIES GENERADAS
'           End If
        Next U&

   End If

End Sub

Private Sub mnuquitar_Click()
       FILX& = GLOBALFRM_FILA&
       If MSFGLOBALFRM.Name <> "MSFSELECCIONADOS" Then
         'solo lo presenta cuando se intenta quitar alguna fila de los ya asignados
         OPX% = COMALTER%("Tenga en Cuenta Que al Desvincular Un Número de Serie a Esta Orden de Fabricación\Deberá Re-Asignar Los Lotes a Los Componentes (sí los Hubiese Cargado)\\Cancelar\Confirmar Desvincular")
         If OPX% < 2 Then
            Exit Sub
         End If
       End If
       '
       If MSFGLOBALFRM.Name = "MSFSELECCIONADOS" Then
         Call cmdEliminaItem(MSFGLOBALFRM, FILX&)
         Exit Sub
       End If
       
       CONDI$ = " TRAZAXSERIE_ID = " & XVALO(MSFGLOBALFRM.TextMatrix(GLOBALFRM_FILA&, 4))
       On Error GoTo ERROR_GUARDAR
       FLEXCONN.BeginTrans
       'BORRO EL CAMPO OF_ARRASTRA QUE ES EL VINCULANTE
       Call ACTUREGISTRO("TRAZAXSERIE", "OF_ARRASTRA", CONDI$, "", REGAF&, "NOMESS")
       NSERIE$ = MSFGLOBALFRM.TextMatrix(GLOBALFRM_FILA&, 3)
       NORFA$ = Label13
       
       'BORRA LA FILA DONDE ESTA EL EQUIPO TRAZABLE -1
       CONDI$ = "TRAZABLE = -1 AND IDSUBOR = '" & NORFA$ & "' AND NUMEROSERIE ='" & NSERIE$ & "'"
       Call BORRAREGISTROS("TRAZAXSERIE", CONDI$, REGAF&, "NOMESS")
       
       'BORRA LAS RESERVAS DE TABLA TRAZAXSERIE
       CONDI$ = "TRAZABLE = 1 AND IDSUBOR = '" & NORFA$ & "' AND NUMEROSERIE ='" & NSERIE$ & "'"
       Call BORRAREGISTROS("TRAZAXSERIE", CONDI$, REGAF&, "NOMESS")
       Call cmdEliminaItem(MSFGLOBALFRM, FILX&)

       
       
ERROR_GUARDAR:
    If Err.Number <> 0 Then
       FLEXCONN.RollbackTrans
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description)
       On Error GoTo 0
       GoTo 99
    Else
      FLEXCONN.CommitTrans
      Call cmdEliminaItem(MSFGLOBALFRM, GLOBALFRM_FILA&)
      Call Label13_Change
    End If
99
    
End Sub
Private Sub mnucopiar_Click()
       Clipboard.Clear
       Clipboard.SetText GLOBALFRM_strContenidoCelda$
End Sub

Private Sub mnuReGenerarReserva_Click()
   On Error Resume Next
   RESP_ZELE_VECT(1) = ""
   On Error GoTo 0
   CAMPOS$ = "CI/A0;Código/A16;Descripción/A28;O.Fabricación/A16;N.Serie/A12;F.Alta;Id/A0"
   Call CARGA_ENCABEZADO(FRMZELECHO.msf2, CAMPOS$, FRMZELECHO)
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, CAMPOS$, FRMZELECHO)
   FRMZELECHO.Caption = ""

   Dim obj As New RECORXET
   Set RS1 = obj.RS_SERIEEYM()
   Call Carga_MSF_Recordset(RS1, CAMPOS$, FRMZELECHO.msf2)
   Set obj = Nothing
   RS1.Close
   Set RS1 = Nothing

   FRMZELECHO.Width = 12500
   FRMZELECHO.Show 1
   If TRIM$(RESP_ZELE_VECT(4)) <> "" Then
      OPX% = COMALTER%("Esta Transacción Vuelve a Leer Las Reservas y Refrescas las Mismas Por N° Serie\Tener en Cuenta Que si Tenía Lotes Asignados Los Mismos Deberá Volvera Asignar\\Cancelar\Regenerar Reservas")
      If OPX% = 2 Then
        CI1% = XVALO(RESP_ZELE_VECT(1))
        OF$ = TRIM$(RESP_ZELE_VECT(4))
        NSERIE$ = TRIM$(RESP_ZELE_VECT(5))
        Call AltaReservaEnTrazaxSerie(OF$, CI1%, 1)
      End If
   End If
   '
   

End Sub

Private Sub MSFSELECCIONADOS_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If Button = 2 Then
     Set MSFGLOBALFRM = MSFSELECCIONADOS
     GLOBALFRM_INDIX% = Index
     REG& = MSFGLOBALFRM.MouseRow
     COL% = MSFGLOBALFRM.MouseCol
     Call MENU_CELDA(MSFGLOBALFRM, REG&, COL%)
  End If
  
End Sub

Private Sub MSFSERIEVINCULADOS_Click()
   REG& = MSFSERIEVINCULADOS.MouseRow
   If REG& < 1 Then Exit Sub
   Call COLOREAR_GRILLA_SOLAFILA(Me.MSFSERIEVINCULADOS, vbYellow, MSFSERIEVINCULADOS.Row, 0, 0, 0)
End Sub

Private Sub MSFSERIEVINCULADOS_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If Button = 2 Then
'      hay que revisar bien para activarla
'     Set MSFGLOBALFRM = MSFSERIEVINCULADOS
'     GLOBALFRM_INDIX% = Index
'     REG& = MSFGLOBALFRM.MouseRow
'     COL% = MSFGLOBALFRM.MouseCol
'     Call MENU_CELDA(MSFGLOBALFRM, REG&, COL%)
  End If

End Sub

Private Sub MSFSRIESDISPONIBLES_DblClick()
    i& = MSFSRIESDISPONIBLES.MouseRow
    J& = MSFSRIESDISPONIBLES.MouseCol
    If i& = 0 Or MSFSRIESDISPONIBLES.Rows <= 1 Then Exit Sub
   If XVALO(MSFSRIESDISPONIBLES.TextMatrix(i&, 4&)) < 1 Then  ' EL ID
     Call COLOREAR_GRILLA_SOLAFILA(Me.MSFSRIESDISPONIBLES, vbYellow, MSFSRIESDISPONIBLES.Row, 0, 1, 0)
     NSERIE$ = MSFSRIESDISPONIBLES.TextMatrix(i&, 3)
     NSERIE$ = TRIM$(InputBox$("Número de Serie", "Ingrese N°. de Serie", NSERIE$))
     MSFSRIESDISPONIBLES.TextMatrix(i&, 3) = NSERIE$
     Call COLOREAR_GRILLA_SOLAFILA(Me.MSFSRIESDISPONIBLES, vbYellow, MSFSRIESDISPONIBLES.Row, 0, 1, 0)
   End If

End Sub
Private Sub MSFSRIESDISPONIBLES_Click()
    i& = MSFSRIESDISPONIBLES.MouseRow
    J& = MSFSRIESDISPONIBLES.MouseCol
    If i& = 0 Or i& > MSFSRIESDISPONIBLES.Rows Then Exit Sub
    Call COLOREAR_GRILLA_SOLAFILA(Me.MSFSRIESDISPONIBLES, vbYellow, MSFSRIESDISPONIBLES.Row, 0, 0, 0)
End Sub


