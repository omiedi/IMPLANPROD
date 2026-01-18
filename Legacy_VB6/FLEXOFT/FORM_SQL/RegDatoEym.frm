VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form RegDatoEym 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Datos"
   ClientHeight    =   9195
   ClientLeft      =   45
   ClientTop       =   915
   ClientWidth     =   14250
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9195
   ScaleWidth      =   14250
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
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
      Height          =   285
      Left            =   1060
      Locked          =   -1  'True
      TabIndex        =   35
      Text            =   " "
      Top             =   1320
      Width           =   360
   End
   Begin VB.TextBox TXTCI 
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
      Height          =   300
      Left            =   8300
      Locked          =   -1  'True
      TabIndex        =   31
      Text            =   " "
      Top             =   0
      Width           =   960
   End
   Begin VB.TextBox Text2 
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
      Height          =   285
      Left            =   10680
      Locked          =   -1  'True
      TabIndex        =   29
      ToolTipText     =   "Doble Click Para Seleccionar Personal de Control"
      Top             =   240
      Width           =   2520
   End
   Begin VB.TextBox TXTMOVIL 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5400
      TabIndex        =   28
      TabStop         =   0   'False
      ToolTipText     =   "Costo segun el proveedor y la via de entrega"
      Top             =   600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text1 
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
      Height          =   285
      Left            =   10680
      Locked          =   -1  'True
      TabIndex        =   4
      ToolTipText     =   "Doble Click Para Seleccionar Personal de Control"
      Top             =   760
      Width           =   2520
   End
   Begin VB.TextBox Text8 
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
      Height          =   300
      Left            =   7575
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   " "
      Top             =   315
      Width           =   1680
   End
   Begin VB.CommandButton Command4 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   9240
      TabIndex        =   1
      Top             =   315
      Width           =   165
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   3960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   870
      Width           =   6450
   End
   Begin VB.TextBox Text11 
      Appearance      =   0  'Flat
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
      Left            =   4920
      TabIndex        =   2
      Text            =   " "
      Top             =   280
      Width           =   945
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
      Height          =   270
      Left            =   5850
      TabIndex        =   3
      Top             =   280
      Width           =   165
   End
   Begin VB.PictureBox Picture13 
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9195
      Left            =   13365
      ScaleHeight     =   9135
      ScaleWidth      =   870
      TabIndex        =   11
      Top             =   0
      Width           =   930
      Begin VB.CommandButton Command23 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   105
         Picture         =   "RegDatoEym.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Visualiza e Imprime Informe de  Fuga a Tierra"
         Top             =   2280
         Width           =   650
      End
      Begin VB.CommandButton Command26 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   105
         Picture         =   "RegDatoEym.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Configuración de Formulario"
         Top             =   1560
         Width           =   650
      End
      Begin VB.CommandButton Command2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   120
         Picture         =   "RegDatoEym.frx":0AAC
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Graba e Imprime Informe"
         Top             =   6720
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   12
         Top             =   7680
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   120
            TabIndex        =   13
            Top             =   120
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command29 
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
         Left            =   105
         Picture         =   "RegDatoEym.frx":0DB6
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Ayuda Flexoft en Línea"
         Top             =   840
         Width           =   650
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "RegDatoEym.frx":10C0
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Salir"
         Top             =   120
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "RegDatoEym.frx":120A
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   7970
         Width           =   1515
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Bindings        =   "RegDatoEym.frx":312C
      Height          =   7360
      Left            =   120
      TabIndex        =   26
      Top             =   1680
      Width           =   13080
      _ExtentX        =   23072
      _ExtentY        =   12991
      _Version        =   393216
      BackColor       =   16777215
      BackColorFixed  =   -2147483645
      ForeColorSel    =   16777215
      AllowBigSelection=   0   'False
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Dispositivo Medidor de Energía"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   720
      TabIndex        =   34
      Top             =   1040
      Width           =   3060
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1440
      TabIndex        =   33
      ToolTipText     =   "Doble Click Para Seleccionar Dispositivo  Medidor de Energía"
      Top             =   1320
      Width           =   2400
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "C.Interno:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6960
      TabIndex        =   32
      Top             =   45
      Width           =   1260
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "O.Fabricación:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9960
      TabIndex        =   30
      Top             =   0
      Width           =   1380
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   27
      Top             =   240
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   9765
      TabIndex        =   25
      ToolTipText     =   "Número de Revisión"
      Top             =   360
      Width           =   360
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9360
      TabIndex        =   24
      Top             =   420
      Width           =   300
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Controló:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9960
      TabIndex        =   23
      Top             =   520
      Width           =   1380
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   10680
      TabIndex        =   22
      ToolTipText     =   "Doble Click Para Seleccionar Personal de Aprobación"
      Top             =   1320
      Width           =   2520
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   120
      Picture         =   "RegDatoEym.frx":3140
      Top             =   1080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "RegDatoEym.frx":344A
      Top             =   1080
      Width           =   480
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4080
      TabIndex        =   21
      Top             =   360
      Width           =   810
   End
   Begin VB.Label Label92 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nº de Serie:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6240
      TabIndex        =   20
      Top             =   360
      Width           =   1260
   End
   Begin VB.Label Label95 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   960
      TabIndex        =   19
      Top             =   135
      Width           =   2880
   End
   Begin VB.Label Label96 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   780
   End
   Begin VB.Label Label88 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   240
      TabIndex        =   17
      Top             =   480
      Width           =   3600
   End
   Begin VB.Label Label24 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Aprobó:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9960
      TabIndex        =   16
      Top             =   1080
      Width           =   1380
   End
   Begin VB.Label Label99 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Comentarios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3480
      TabIndex        =   15
      Top             =   600
      Width           =   1530
   End
   Begin VB.Label Label4 
      BackColor       =   &H00CDDADA&
      Caption         =   "Registro de Datos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4200
      TabIndex        =   14
      Top             =   0
      Width           =   2175
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo "
   End
   Begin VB.Menu mnuTabAux 
      Caption         =   "Tablas Auxiliares"
      Begin VB.Menu mnuCarTabAux 
         Caption         =   "Carga de Tablas Auxiliares"
      End
   End
End
Attribute VB_Name = "RegDatoEym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TXTRAZA_COFUTIE$

Sub Actualizar_Validaciones_En_Grilla()
   For i& = 1 To MSF.Rows - 1
      MSF.TextMatrix(i&, 0) = VALIDARRESULTADO$(MSF, i&, 6)
   Next i&
End Sub

Private Sub Command1_Click()
   '\\\OT-08-0936-OD-08/10/08///
    Call CIERRARCH("*.*")
    Unload Me
    '\\\OT-08-0936-OD-FIN///
End Sub
Private Sub Command2_Click()
   '\\\OT-08-0936-OD-08/10/08///
    Command2.Enabled = False
    '
    'VALIDACION POR FECHA
   
    VDEVV$ = TRIM$(Text11.TEXT)
    If FECHANUM(VDEVV$) < 1 Then
      Call MENSERR(24, "Fecha Inválida")
      GoTo 99
    End If
    '
    'VALIDACION DE NUMERO DE SERIE
    '
    NUSERIE$ = TRIM$(Text8)
    CANTRE& = CantRegistros&("TANUMSE", "NumeroSerie ='" & NUSERIE$ & "'")
    If CANTRE& < 1 Then
       Call MENSERR(24, "Número de Serie Incorrecto")
       GoTo 99
    End If
    '
    '
    If Label8 = "" Then
      OPX1% = COMALTER%("Falta Informar Quién Aprobó\\Cancelar\Continuar")
      If OPX1% < 2 Then GoTo 99
    End If
    '
    If Label12 = "" Then
      OPX1% = COMALTER%("Falta Informar Dispositivo Medidor de Energía\\Cancelar\Continuar")
      If OPX1% < 2 Then GoTo 99
    End If

    If TRIM$(Text1) = "" Then
      Call MENSERR(24, "Falta Informar Quién Controló")
      GoTo 99
    End If

    Screen.MousePointer = 11
    '

    ULTREVIS% = XVALO(VALOBADA("REGIDATO", "MAX(REVISION)", "NumeroSerie ='" & NUSERIE$ & "'"))
    REVIS% = ULTREVIS% + 1

    CIXI% = CODINT%(Label95)
    RESULTADOX$ = "C"
    
    EXCEP$ = TRIM$(EXCEPCION_CONTROL_EYM$(CIXI%, "REGIDATO"))
    Call TEXTOLOTES(EXCEP$, ";")
    '
    'ACA VALIDO LOS DATOS LEIDOS CONSIDERANDO LAS EXCEPCIONES CONFIGURADAS
    'SALVO EXCEPCIONES DEBE ESTAR TODO OK PARA QUE ESTE CONFORME
    'CASO CONTRARIO EL RESULDATO ES NC
    If EXCEP$ <> "" Then
        For i& = 1 To MSF.Rows - 1
            On Error Resume Next
            'VALIDA SI TIENE EXCEPCIONES
            AA = UBound(CADENATEX$)
            J& = 0
            While J& < UBound(CADENATEX$)
              J& = J& + 1
              If i& = XVALO(CADENATEX$(J&)) Then
                 GoTo 20
              End If
            Wend
            On Error GoTo 0
            If TRIM$(UCase$(MSF.TextMatrix(i&, 8))) <> "OK" Then
               RESULTADOX$ = "NC"
               Exit For
            End If
20      Next i&
    Else
        For i& = 1 To MSF.Rows - 1
            If TRIM$(UCase$(MSF.TextMatrix(i&, 0))) <> "OK" Then
'               Label5.Caption = Chr(254)
'               Label12.Caption = Chr(168)
'               Label5.Refresh
'               Label12.Refresh
               RESULTADOX$ = "NC"
               Exit For
            End If
         Next i&
    End If
    On Error GoTo ERROR_GUARDAR
    FLEXCONN.BeginTrans

    SQLX$ = "SELECT * FROM REGIDATO"
    SQLX$ = SQLX$ + " WHERE NUMEROSERIE = '" & Text8 & "'"
    SQLX$ = SQLX$ + " AND REVISION = " & REVIS%
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    With rs
       For i& = 1 To MSF.Rows - 1
            .AddNew   ' YA CONVERTIDO
            On Error Resume Next
            !CodiEmpr = CodiEmp%
            On Error GoTo 0
            !NumeroSerie = TRIM$(Text8)
            !Modelo = CIXI%
            !OP = XVALO(MSF.TextMatrix(i&, 1)) * 10
            !Descripcion = MSF.TextMatrix(i&, 2)
            !Rango_Acep_Min = MSF.TextMatrix(i&, 3)
            !VM = MSF.TextMatrix(i&, 4)
            !Rango_Acep_Max = MSF.TextMatrix(i&, 5)
            !VALOR_RESUL = MSF.TextMatrix(i&, 6)
            !Unimed = MSF.TextMatrix(i&, 7)
            !Bloqueado = "SI"
            !Revision = REVIS%
            !Controlo = Left$(Text1, 32)
            !Comentarios = Left$(Text9, 512)
            !APROBO = Left$(Label8, 32)
            !IDSUBOR = Left$(Text2, 16)
            !Fech_Ensa = CVDAT(FECHANUM(Text11))
            !Resultado = RESULTADOX$
            !DISPOMEDINUM = XVALO(Text3)
            !DISPOMEDIX = Label12
            .Update
       Next i&
            
    End With
    '
    '
    'GRABACION EN EL ANOTADOR (HISTORIAL DE EQUIPOS)
    CUERPO$ = "REGISTRO DE DATOS" & vbCrLf
    CUERPO$ = CUERPO$ + "Controló: " & UCase$(TRIM$(Text1))
    CUERPO$ = CUERPO$ + " - Aprobó: " & UCase$(TRIM$(Label8))
    Call ANOTANUSE(NUSERIE$, CUERPO$)
    '
ERROR_GUARDAR:
    If Err.Number <> 0 Then
       FLEXCONN.RollbackTrans
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description)
       rs.Close
       Set rs = Nothing
       On Error GoTo 0
    Else
      FLEXCONN.CommitTrans
      Call COMUNI("Proceso Completo")
      rs.Close
      Set rs = Nothing
      On Error GoTo 0
    End If

    Call PRINTREGIDATO(NUSERIE$, REVIS%)
    '
    Image1.Visible = True
    Image3.Visible = False
    Command23.Enabled = True
    '
    Call BLANFORMU
    '
    Screen.MousePointer = 1
    '
99  Command2.Enabled = True
   '\\\OT-08-0936-OD-fin///
End Sub
Sub PRINTREGIDATO(NUSERIE$, ULTREVIS%)
        'SELECCION DE OF DE LOS PRODUCTOS TERMINADOS
        'MODIFICADO EL 23/01/18
        NORFA$ = TRIM$(Text2)
        If NORFA$ = "" Then
           Call COMUNI("No Hay una Orden de Fabricación Activa")
           Exit Sub
        End If
        '
        NSERIE$ = TRIM$(Text8)
        If NSERIE$ = "" Then
           Call COMUNI("Falta Seleccionar Un Número de Serie")
           Exit Sub
        End If
        '
        'VALIDA SI YA SEW GRABARON LOS DATOS DEL CPU
        CI1% = XVALO(Me.TXTCI)
        '
        Unload VENTABNEW
        Campos$ = "DATO1/A16;DATO2/A16;DATO3/A16;DATO4/A16;DATO1/A16;DATO6/A16"
        VENTANA.Campox = Campos$

        Call CopiarMsfAOtroMsf(Me.MSF, VENTABNEW.MSF)
        VENTANA.Inicializar = 0
        VENTANA.TITULO = "Registro de Datos"
        VENTANA.Inicializar = 1
        VENTABNEW.Height = 7500
        VENTABNEW.MSF.Font = "Roboto Mono"
        VENTABNEW.MSF.Font.Size = 8
        k& = MSF.Rows - 1
        For i& = 1 To 9
          VENTABNEW.MSF.Rows = k& + i& + 1
          Select Case i&
             Case 1: VENTABNEW.MSF.TextMatrix(k& + i, 5) = Text1  'CONTROLO
             Case 2: VENTABNEW.MSF.TextMatrix(k& + i, 5) = Label8   'APROBO
             Case 3: VENTABNEW.MSF.TextMatrix(k& + i, 5) = Label95  'CODIGO
             Case 4: VENTABNEW.MSF.TextMatrix(k& + i, 5) = Label88  'DESCRIPCION
             Case 5: VENTABNEW.MSF.TextMatrix(k& + i, 5) = Text11  'FECHA
             Case 6: VENTABNEW.MSF.TextMatrix(k& + i, 5) = NSERIE$  'N.SERIE
             Case 7: VENTABNEW.MSF.TextMatrix(k& + i, 5) = NORFA$  'NORFA$
             Case 8: VENTABNEW.MSF.TextMatrix(k& + i, 5) = Label12  'NORFA$
             Case 9: VENTABNEW.MSF.TextMatrix(k& + i, 5) = Text9  'COMENTARIOS
          End Select
        Next i&
        '
        ' EL EXCEL SE MODIFICO EL MISMO TIENE 2 REGISTROS MENOS SE CONSIDERA AHORA DEBEN
        ' CONVIVIR 2 FORMAS DE PRESENTAR E IMPRIIMIR LOS CONTROLES EL NUEVO Y EL VIEJO
        ' LOS DOS QUEDARON CONFIGURABLES.
        
        If MSF.Rows >= 47 Then
          RUTAEXCEL$ = Control("REGIDATO", "ARCHDATO")
        ElseIf MSF.Rows <= 45 Then  'NUEVO FORMULARIO
          RUTAEXCEL$ = Control("REGIDATO", "ARCHDAT2")
        End If
        If TRIM$(RUTAEXCEL$) = "" Then
           Call COMUNI("Falta Configurar Ruta de Archivo de  Impresión")
           GoTo 90
        End If
        PDES$ = Control$(IDENTER$, "PRIREGDA")
        If PDES$ = "" Then PDES$ = Control$("*", "PRIREGDA")
        Call Param_Print_Excel(RUTAEXCEL$, "094", VENTABNEW.MSF, PDES$, "5", 10, 1, "", 1)
90      Unload VENTABNEW


End Sub

Private Sub Command23_Click()
   NUSERIE$ = TRIM$(Text8)
   CANTRE& = CantRegistros&("TANUMSE", "NumeroSerie ='" & NUSERIE$ & "'")
   If CANTRE& < 1 Then
      Call MENSERR(24, "Número de Serie Incorrecto")
      Exit Sub
   End If
'   Call PRINTMEDI(0) '  0 IMPRESION SIMPLE --- UNO GRABACION
    ULTREVIS% = XVALO(Label13)
    CIXI% = CODINT%(TRIM$(Label95))
    Call PRINTREGIDATO(NUSERIE$, ULTREVIS%)
    '
End Sub
Private Sub Command26_Click()
   Command26.Enabled = False
   FSREGDAT.Show 1
   Command26.Enabled = True
End Sub
Private Sub Command4_Click()
   '\\\OT-08-0936-OD-08/10/08///
   Call Text8_DblClick
   '\\\OT-08-0936-OD-FIN///
End Sub


Private Sub Command6_Click()
   '\\\OT-08-0936-OD-08/10/08///
    If Image3.Visible = False Then
      Call CommSel("SELFECHA", Command6, Text11)
    Else
      Command4.SetFocus
    End If
    '\\\OT-08-0936-OD-FIN///
End Sub


Private Sub Form_Load()
    '\\\OT-08-0936-OD-06/10/08///
    Call ABREORFAB
    Call BLANFORMU
    FF% = HOY(HO$)
    Text11.TEXT = HO$
    MSF.Rows = 1
    Call CARGA_REGISTRO_DATOS
    '\\\OT-08-0936-OD-FIN///
End Sub
Private Sub Frame5_DragDrop(Source As Control, X As Single, Y As Single)
End Sub

Private Sub Label2_Click()
    If Label2.Caption = Chr(254) Then
       Label2.Caption = Chr(168)
    Else
       Label2.Caption = Chr(254)
       Label5.Caption = Chr(168)
    End If

End Sub

Private Sub Label5_Click()
    If Label5.Caption = Chr(254) Then
       Label5.Caption = Chr(168)
    Else
       Label5.Caption = Chr(254)
       Label2.Caption = Chr(168)
    End If
End Sub
Private Sub Image1_DblClick()
   If DERACCE("REGIDATO", "Modificacion  de Registros de Datos") > 1 Then
        Image1.Visible = False
        Image3.Visible = True
        Command2.Enabled = False
        Command23.Enabled = True '\\\OT-08-1064-OD-21/11/08-MEDITEA- HABILITA LA IMPRESORA
        If TRIM$(Text8) <> "" Then
          'PONGO LA FECHA DEL ULTIMO INFORME
          NUSERIE$ = TRIM$(Text8)
          CONDI$ = "NumeroSerie ='" & NUSERIE$ & "'"
          ULTIMOENSA% = XVALO(VALOBADA("REGIDATO ", "MAX(REVISION)", CONDI$))
          CONDI$ = CONDI$ + " AND REVISION = " & ULTIMOENSA%
          Text11 = FECHATEX(CVINT(VALOBADA("REGIDATO", "Fech_Ensa", CONDI$, "NOMESS")))
        End If
   End If
End Sub

Private Sub Image3_DblClick()
   If DERACCE("REGIDATO", "Modificacion  de Registro de Datos") > 1 Then
        Image1.Visible = True
        Image3.Visible = False
        Command2.Enabled = True
        Command23.Enabled = False
        'PONGO FECHA DEL DIA CUANDO VA A REALIZAR OTRA REVISION
        HOII% = HOY(HOS$)
        Text11 = FECHATEX$(HOII%)
   End If
End Sub

Private Sub Label12_DblClick()
    If Image3.Visible = True Then Exit Sub ' SI ESTA BLOQUEADO NO CAMBIA
    Campos$ = "Código/f6;Disp.Medidor Energía/A24"
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    FRMZELECHO.Caption = "TABLA PERSONAL CONTROL"
    
    Dim obj As New RECORXET
    Set rs1 = obj.RS_DISPOMEDI_EYM()
    Call Carga_MSF_Recordset(rs1, Campos$, FRMZELECHO.msf2, 1)
    On Error Resume Next
    Set obj = Nothing
    rs1.Close
    Set rs1 = Nothing
    On Error GoTo 0
        
    FRMZELECHO.Width = 8000
50  FRMZELECHO.Show 1

    Text3 = TRIM(RESP_ZELE_VECT(1))
    Label12 = TRIM(RESP_ZELE_VECT(2))
   '
End Sub

Private Sub Label8_Click()
    Campos$ = "Código/f6;Nombre/A32"
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    FRMZELECHO.Caption = "TABLA PERSONAL APROBACIÓN"
    
    Dim obj As New RECORXET
    Set rs1 = obj.RS_PERSOAPRO_EYM()
    Call Carga_MSF_Recordset(rs1, Campos$, FRMZELECHO.msf2, 1)
    On Error Resume Next
    Set obj = Nothing
    rs1.Close
    Set rs1 = Nothing
    On Error GoTo 0
        
    FRMZELECHO.Width = 8000
50  FRMZELECHO.Show 1

    Label8 = TRIM(RESP_ZELE_VECT(2))

End Sub

Private Sub Label95_Change()
    CIXI% = CODINT(Label95)
    If CIXI% > 0 Then
        EXCEP$ = TRIM$(EXCEPCION_CONTROL_EYM$(CIXI%, "CONTFINEYM"))
        If EXCEP$ = "" Then Exit Sub
        If InStr(EXCEP$, ";") < 1 Then Exit Sub
        Call TEXTOLOTES(EXCEP$, ";")
        '
        For i& = 1 To MSF.Rows - 1
            On Error Resume Next
            'VALIDA SI TIENE EXCEPCIONES
            
            X1 = UBound(CADENATEX$)
            While J& < UBound(CADENATEX$)
              J& = J& + 1
              If i& = XVALO(CADENATEX$(J&)) Then
                 MSF.TextMatrix(i&, 0) = "N/A"
                 MSF.TextMatrix(i&, 6) = "N/A"
              End If
            Wend
        Next i&
    End If

End Sub

Private Sub mnuCarTabAux_Click()
   Call GENERAR_ESTRUCTURAS_SQL
   Call PRESENTA_TABLAS_AUXILIARES
   
   
End Sub

Private Sub MSF_Click()
    
    If Image3.Visible = True Then Exit Sub ' SI ESTA BLOQUEADO NO CAMBIA
    REG& = MSF.MouseRow
    COLUM% = MSF.MouseCol
    If REG& = 0 Or REG& > MSF.Rows - 1 Then Exit Sub
    If MSF.TextMatrix(REG&, 0) = "N/A" Then Exit Sub
    
    If COLUM% <> 6 Then
      TXTMOVIL.Visible = False
      Exit Sub
    End If
    '
    IDFILA& = XVALO(MSF.TextMatrix(REG&, 1))
    'If (IDFILA& >= 6 And IDFILA& <= 17) Or (IDFILA& = 23) Or (IDFILA& >= 37 And IDFILA& <= 42) Then
    If (IDFILA& >= 6 And IDFILA& <= 15) Or (IDFILA& = 21) Or (IDFILA& >= 35 And IDFILA& <= 40) Then
        MSF.col = 6
        With TXTMOVIL
          .ZOrder 0
          .Width = MSF.CellWidth
          .Left = MSF.CellLeft + MSF.Left
          .Top = MSF.CellTop + MSF.Top
          .Height = MSF.CellHeight
          .Visible = True
          .ToolTipText = MSF.row
          .TEXT = MSF.TEXT
          .SetFocus
        End With
    Else
        
        VALOX1$ = TRIM$(MSF.TextMatrix(REG&, COLUM%))
        If VALOX1$ = "" Then
           VALOX1$ = "OK"
           Call COLOREAR_TEXTO_SOLO_UNA_AFILA(MSF, vbBlack, REG&, COLUM%)
        ElseIf VALOX1$ = "OK" Then
           VALOX1$ = "MAL"
           Call COLOREAR_TEXTO_SOLO_UNA_AFILA(MSF, vbRed, REG&, COLUM%)
        ElseIf VALOX1$ = "MAL" Then
           VALOX1$ = ""
        End If
        MSF.TextMatrix(REG&, COLUM%) = VALOX1$
        MSF.TextMatrix(REG&, 0) = VALOX1$
    End If
    
End Sub



Private Sub Timer1_Timer()

End Sub

Private Sub Text1_DblClick()
    If Image3.Visible = True Then Exit Sub ' SI ESTA BLOQUEADO NO CAMBIA
    Campos$ = "Código/f6;Nombre/A32"
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    FRMZELECHO.Caption = "TABLA PERSONAL CONTROL"
    
    Dim obj As New RECORXET
    Set rs1 = obj.RS_PERSOCONTROL_EYM()
    Call Carga_MSF_Recordset(rs1, Campos$, FRMZELECHO.msf2, 1)
    On Error Resume Next
    Set obj = Nothing
    rs1.Close
    Set rs1 = Nothing
    On Error GoTo 0
        
    FRMZELECHO.Width = 8000
50  FRMZELECHO.Show 1

    Text1 = TRIM(RESP_ZELE_VECT(2))

End Sub

Private Sub Text9_Click()
With MSF
For i = 0 To .Cols - 1
   .ColAlignment(i) = 4
Next i
End With
End Sub

Private Sub TXTMOVIL_KeyDown(KeyCode As Integer, Shift As Integer)
'    '38 y 40
'
'0   If KeyCode = 40 Then 'ABAJO
'      REG& = XVALO(TXTMOVIL.ToolTipText)
'      MSF.TextMatrix(REG&, 6) = TXTMOVIL
'      If REG& >= MSF.Rows - 1 Then
'         TXTMOVIL.ToolTipText = 1
'
'         GoSub POSICIONA
'      Else
'
'          'GoSub POSICIONA
'          TXTMOVIL_KeyPress (13)
'      End If
'    End If
'    If KeyCode = 38 Then
'      REG& = XVALO(TXTMOVIL.ToolTipText) - 1
'      MSF.TextMatrix(REG&, 6) = TXTMOVIL
'      If REG& < 1 Then
'         TXTMOVIL.ToolTipText = MSF.Rows - 1
'
'         GoSub POSICIONA
'      Else
'          TXTMOVIL.ToolTipText = REG&
'          GoSub POSICIONA
'      End If
'    End If
'    Exit Sub
'
'POSICIONA:
'              REG& = XVALO(TXTMOVIL.ToolTipText)
'              MSF.Row = REG&
'                 With TXTMOVIL
'                  .ToolTipText = REG&
'                  .ZOrder 0
'                  'recordar que si un campo esta oculto el mismo no debe estar editable
'                   .TEXT = TRIM$(MSF.TEXT)
'                  On Error Resume Next
'                  .Width = MSF.CellWidth
'                  On Error GoTo 0
'                  .Left = MSF.CellLeft + MSF.Left
'                  .Top = MSF.CellTop + MSF.Top
'                  .Visible = True
'                  Call PINTATEXT(TXTMOVIL)
'                End With

End Sub

 Private Sub TXTMOVIL_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then TXTMOVIL.Visible = False: Exit Sub
    COLU1& = 6
    MSF.col = COLU1&
    
    'RUTINA PARA QUE AL DAR ENTRER O ESCAPE SAQUE EL FOCO Y PONGA COMO NO VISIBLE EL TEXT DE EDICION
    If KeyAscii = 13 Then
      REG& = XVALO(TXTMOVIL.ToolTipText)
      MSF.row = REG&
'      If TRIM$(MSF.TextMatrix(REG&, 3)) = "" And TRIM$(MSF.TextMatrix(REG&, 3)) = "" Then
      IDFILA& = XVALO(MSF.TextMatrix(REG&, 1))
      If (IDFILA& >= 6 And IDFILA& <= 15) Or (IDFILA& = 21) Or (IDFILA& >= 35 And IDFILA& <= 40) Then
         MSF.TEXT = TRIM$(FORMATNUM$(XVALO(TXTMOVIL.TEXT), "F12.2"))
         REG& = XVALO(TXTMOVIL.ToolTipText)
         MSF.TextMatrix(REG&, 0) = VALIDARRESULTADO$(MSF, REG&, 6)
      Else
         GoTo 20
         MSF.TEXT = TRIM$(TXTMOVIL.TEXT)
         REG& = XVALO(TXTMOVIL.ToolTipText)
         MSF.TextMatrix(REG&, 0) = VALIDARRESULTADO$(MSF, REG&, 6)
      End If
20    TXTMOVIL.Visible = False
      On Error Resume Next
'      REG& = XVALO(CADENATEX$(1))
      On Error GoTo 0
      If REG& >= MSF.Rows - 1 Then
        Exit Sub
      Else
         MSF.row = REG& + 1
         MSF.col = 6
         With TXTMOVIL
          .ToolTipText = MSF.row
          .ZOrder 0
          'recordar que si un campo esta oculto el mismo no debe estar editable
          On Error Resume Next
          .Width = MSF.CellWidth
          On Error GoTo 0
          .Left = MSF.CellLeft + MSF.Left
          .Top = MSF.CellTop + MSF.Top
           .TEXT = TRIM$(MSF.TEXT)
          .Visible = True
          Call PINTATEXT(TXTMOVIL)
        End With
      'Call ACTUCANTI
      End If
    End If
    

End Sub



Private Sub MSF_MouseUp(Button As Integer, _
         Shift As Integer, X As Single, Y As Single)
'          Static CurrentWidth As Single
'          If MSF.Rows <= 1 Then Exit Sub
'          If MSF.CellWidth <> CurrentWidth Then
'              On Error Resume Next
'              TXTMOVIL.Width = MSF.CellWidth
'              CurrentWidth = MSF.CellWidth
'              On Error GoTo 0
'          End If
End Sub
Private Sub TXTMOVIL_Change()
   REG& = XVALO(TXTMOVIL.ToolTipText)
   
End Sub

Private Sub TXTMOVIL_GotFocus()
     Call PINTATEXT(TXTMOVIL)
End Sub

Private Sub MSF_Scroll()
   If TXTMOVIL.Visible = True Then
     'TXTMOVIL.Left = MSF1.CellLeft
     REG& = XVALO(TXTMOVIL.ToolTipText)
     MSF.TEXT = TXTMOVIL
     TXTMOVIL.Visible = False
   End If
End Sub



'     COLU1& = 6
'     If TXTMOVIL.Visible = True Then
'       REG& = XVALO(TXTMOVIL.ToolTipText)
'       MSF.TextMatrix(REG&, 6) = TRIM$(FORMATNUM$(XVALO(MSF.TextMatrix(REG&, 6)), "F12.2"))
''       TXTMOVIL.Visible = False
'     End If
'     Call ACTU_TOTAL



Private Sub Text1_GotFocus()
  If Image3.Visible = True Then Command4.SetFocus '\\\OT-08-1066-OD-21/11/08-MEDITEA QUE NO SE PUEDA EDITAR SI TIENE EL CANDADO VISIBLE
End Sub




Private Sub Text7_GotFocus()
   If Image3.Visible = True Then Command4.SetFocus
End Sub
Private Sub Text8_Change()
   
   '\\\OT-08-0936-OD-08/10/08///
   Command2.Enabled = True
   Image3.Visible = False
   Image1.Visible = True
   Image1.Enabled = False
   TXTCI = ""
   Command23.Enabled = True
   TXTRAZA_COFUTIE$ = ""
   Call LIMPIAINFORMACION
   
   '\\\OT-08-0936-OD-08/10/08///
End Sub
Private Sub Text8_DblClick()
   '\\\OT-08-0936-OD-08/10/08///
    
    NSERIE1$ = SERIES_REGISTRODEDATOS
    If NSERIE1$ <> "" Then
       Text8 = ""

       Text8 = NSERIE1$
      
       Label95 = TRIM$(RESP_ZELE_VECT(2))
       TXTCI = CODINT%(Label95)
       
       Label88 = TRIM$(RESP_ZELE_VECT(3))
       Text2 = TRIM$(RESP_ZELE_VECT(4))
       Text11 = TRIM$(RESP_ZELE_VECT(5))
       Call Text8_KeyPress(13)
       Call Actualizar_Validaciones_En_Grilla
    End If
 
   '\\\OT-08-0936-OD-08/10/08///
End Sub
 Function SERIES_REGISTRODEDATOS()

   On Error Resume Next
   RESP_ZELE_VECT(1) = ""
   On Error GoTo 0
   Campos$ = "N°Serie/A12;Código/A16;Descripción/A28;O.Fabric./A12;F.Alta/D8"
   Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO)
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
   FRMZELECHO.Caption = ""
 
   Dim obj As New RECORXET
   Set rs1 = obj.RS_COD_LLEVAN_CONTROL("REGIDATO")
   Call Carga_MSF_Recordset(rs1, Campos$, FRMZELECHO.msf2)
   Set obj = Nothing
   rs1.Close
   Set rs1 = Nothing

   FRMZELECHO.Width = 10500
   FRMZELECHO.Show 1
   If TRIM$(RESP_ZELE_VECT(1)) <> "" Then
      SERIES_REGISTRODEDATOS = TRIM$(RESP_ZELE_VECT(1))
   End If
    '
End Function

Sub BLANFORMU()
    Text1 = "" '\\\OT-08-1065-OD-21/11/08-MEIDTEA-LIMPIA EL TEXT DEL TECNIO
    Text7 = ""
    Text8 = ""
    Text9 = ""
    Label13 = "" '\\\OT-08-1064-OD-21/11/08-MEDITEA-LIMPIO EL LABEL DE REVISION
    Label88 = ""
    Label95 = ""
    Text11 = FECHATEX$(HOY(HO$))
    Label12 = ""
'    Label98 = "I354"
    RIGIDIEL07.Label98 = ""
    RIGIDIEL07.Text5 = ""
    
End Sub
Sub LIMPIAINFORMACION()
    Text1 = "" '\\\OT-08-1065-OD-21/11/08-MEIDTEA-LIMPIA EL TEXT DEL TECNIO
    Text7 = ""
    Text9 = ""
    Label13 = "" '\\\OT-08-1064-OD-21/11/08-MEDITEA-LIMPIO EL LABEL DE REVISION
    Label88 = ""
    Label95 = ""
    Label98 = ""
    'COMO SE CAMBIO LA ESTRUCTURA , SI YA ESTA GRABADA LA VERSION LA MUESTRA COMPLETA
    'EN CASO QUE SEA UN CONTROL NUEVO LLAMA A CARGA_REGISTRO_DATOS PARA QUE GENERE LA GRILLA C/NEW ESTRUCTURA.
    CARGA_REGISTRO_DATOS

'    Call LIMPIAR_MSF_1_COLUMNA_SIN_BORRAR_REGISTROS(MSF, 6)
'    Call LIMPIAR_MSF_1_COLUMNA_SIN_BORRAR_REGISTROS(MSF, 0)
    Call VOLVER_ABLANCO_TODA_LA_GRILLA(MSF)
End Sub

Private Sub Text8_GotFocus()
'  Call Command4.SetFocus
'  Call Command4_Click
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
   
   If KeyAscii = 13 Then
             
             NUSERIE$ = TRIM$(Text8)

'             CANTRE& = CantRegistros&("TANUMSE", "NumeroSerie ='" & NUSERIE$ & "'")
'             If CANTRE& < 1 Then
'               Call MENSERR(24, "Número de Serie Incorrecto")
'               Exit Sub
'             End If
             '
             Screen.MousePointer = 11
             
'             CI% = XVALO(VALOBADA("TANUMSE", "Cod_Inte", "NumeroSerie ='" & NUSERIE$ & "'"))
'             If CI% > 0 Then
'              Label95 = CODEXT$(CI%)
'              Label88 = DESCRIT0$(CI%)
'             End If
'             Text2 = VALOBADA("TANUMSE", "IDSUBOR", "NumeroSerie ='" & NUSERIE$ & "'")
             
             ULTREVIS% = XVALO(VALOBADA("REGIDATO", "MAX(REVISION)", "NumeroSerie ='" & NUSERIE$ & "' AND MODELO = " & XVALO(TXTCI)))
             Label13 = ULTREVIS%
             
             
             BLQUEO$ = "": VUELTA% = 0
             SQLX$ = "SELECT  *  FROM REGIDATO WITH(NOLOCK)"
             SQLX$ = SQLX$ + " WHERE NumeroSerie ='" & NUSERIE$ & "'"
             SQLX$ = SQLX$ + " AND REVISION = " & ULTREVIS% & ""
             SQLX$ = SQLX$ + " order by OP "
             Set rs = READSET(SQLX$)
             With rs
               Do While Not .EOF
                      i& = i& + 1
                      MSF.Rows = i& + 1
                      MSF.TextMatrix(i&, 1) = XVALO(!OP) / 10
                      MSF.TextMatrix(i&, 2) = SAFETEXT$(!Descripcion)
                      MSF.TextMatrix(i&, 3) = SAFETEXT$(!Rango_Acep_Min)
                      MSF.TextMatrix(i&, 4) = SAFETEXT$(!VM)
                      MSF.TextMatrix(i&, 5) = SAFETEXT$(!Rango_Acep_Max)
                      MSF.TextMatrix(i&, 6) = SAFETEXT$(!VALOR_RESUL)
                      MSF.TextMatrix(i&, 7) = SAFETEXT$(!Unimed)
                      
                      '
                      If VUELTA% < 1 Then
                        Text11 = FECHATEX$(CVINT(!Fech_Ensa))
                        BLQUEO$ = SAFETEXT$(!Bloqueado)
                        Text1 = SAFETEXT$(!Controlo)
                        Text9 = SAFETEXT$(!Comentarios)
                        Label8 = SAFETEXT$(!APROBO)
                        VL% = 168: VLL% = 168
                        VL1$ = TRIM$(SAFETEXT$(!DP006)): If VL1$ = "." Then VL% = 254
                        Label5 = Chr$(VL%)
                        VL1$ = TRIM$(SAFETEXT$(!dp100)): If VL1$ = "." Then VLL% = 254
                        Label2 = Chr$(VLL%)
                        Text3 = XVALO(!DISPOMEDINUM)
                        Label12 = SAFETEXT$(!DISPOMEDIX)

                        VUELTA% = 1
                      End If
                      '
                      .MoveNext
                  Loop
            End With
        rs.Close
        Set rs = Nothing
        Screen.MousePointer = 1

        
        If BLQUEO$ = "SI" Then
                Image3.Visible = True
                Image1.Visible = False
                Image1.Enabled = True
                Command2.Enabled = False
        Else
                Command23.Enabled = False '\\\OT-08-1064-OD-21/11/08-MEIDTEA BLOQUEA LA IMPRESION
                Text11 = FECHATEX$(HOY(HO$)) 'SI NO TIENE CONTROLES LE APLICO LA FECHA DEL DIA
        End If
        '
        Call Actualizar_Validaciones_En_Grilla
        
    End If
End Sub

Private Sub Text9_GotFocus()
   If Image3.Visible = True Then Command4.SetFocus
End Sub

Sub CARGA_REGISTRO_DATOS()
    Call MSF.Clear
    Campos$ = "Op./F5;Descripción/A60;Rango aceptable/F4.2;Rango aceptable/F12.2;Rango aceptable/F4.2;Registro/F8.2;U./A4"
    Call CARGA_ENCABEZADO(MSF, Campos$, Me, , , , 7)
    
    MSF.MergeRow(0) = True
    MSF.MergeCells = flexMergeFree
    MSF.Rows = 45
    MSF.FixedRows = 1

    MSF.TextMatrix(1, 1) = "1"
    MSF.TextMatrix(1, 2) = "Ensambles y conexionados"
    
    MSF.TextMatrix(2, 1) = "2"
    MSF.TextMatrix(2, 2) = "Circuito de red de corriente alterna"
    
    MSF.TextMatrix(3, 1) = "3"
    MSF.TextMatrix(3, 2) = "Configuración EEPROM"
    
    MSF.TextMatrix(4, 1) = "4"
    MSF.TextMatrix(4, 2) = "Audio e indicadores luminosos"
    
    MSF.TextMatrix(5, 1) = "5"
    MSF.TextMatrix(5, 2) = "Fuente de alimentación general"
    MSF.TextMatrix(5, 3) = "14.70"
    MSF.TextMatrix(5, 4) = "< VM >"
    MSF.TextMatrix(5, 5) = "15.30"
    MSF.TextMatrix(5, 7) = "V"
    
    MSF.TextMatrix(6, 1) = "6"
    MSF.TextMatrix(6, 2) = "Fuente de alimentación +6,8V"
    MSF.TextMatrix(6, 3) = "6.50"
    MSF.TextMatrix(6, 4) = "< VM >"
    MSF.TextMatrix(6, 5) = "7.00"
    MSF.TextMatrix(6, 7) = "V"
    
    MSF.TextMatrix(7, 1) = "7"
    MSF.TextMatrix(7, 2) = "Fuente de alimentación +5V ECG"
    MSF.TextMatrix(7, 3) = "4.95"
    MSF.TextMatrix(7, 4) = "< VM >"
    MSF.TextMatrix(7, 5) = "5.05"
    MSF.TextMatrix(7, 7) = "V"
    
    MSF.TextMatrix(8, 1) = "8"
    MSF.TextMatrix(8, 2) = "Fuente de alimentación +5V CPU"
    MSF.TextMatrix(8, 3) = "4.95"
    MSF.TextMatrix(8, 4) = "< VM >"
    MSF.TextMatrix(8, 5) = "5.05"
    MSF.TextMatrix(8, 7) = "V"
    
'    MSF.TextMatrix(9, 1) = "9"
'    MSF.TextMatrix(9, 2) = "Fuente de alimentación +12V ECG"
'    MSF.TextMatrix(9, 3) = "11.0"
'    MSF.TextMatrix(9, 4) = "< VM >"
'    MSF.TextMatrix(9, 5) = "13.0"
'    MSF.TextMatrix(9, 7) = "V"
'
'    MSF.TextMatrix(10, 1) = "10"
'    MSF.TextMatrix(10, 2) = "Fuente de alimentación -12V ECG"
'    MSF.TextMatrix(10, 3) = "-13.0"
'    MSF.TextMatrix(10, 4) = "< VM >"
'    MSF.TextMatrix(10, 5) = "-11.0"
'    MSF.TextMatrix(10, 7) = "V"
    'a partir de aqui desconte 2
    
    MSF.TextMatrix(9, 1) = "9"
    MSF.TextMatrix(9, 2) = "Fuente de alimentación +8V ECG"
    MSF.TextMatrix(9, 3) = "7.70"
    MSF.TextMatrix(9, 4) = "< VM >"
    MSF.TextMatrix(9, 5) = "8.30"
    MSF.TextMatrix(9, 7) = "V"
    
    MSF.TextMatrix(10, 1) = "10"
    MSF.TextMatrix(10, 2) = "Fuente de alimentación -8V ECG"
    MSF.TextMatrix(10, 3) = "-8.30"
    MSF.TextMatrix(10, 4) = "< VM >"
    MSF.TextMatrix(10, 5) = "-7.70"
    MSF.TextMatrix(10, 7) = "V"
    
    MSF.TextMatrix(11, 1) = "11"
    MSF.TextMatrix(11, 2) = "Fuente de alimentación +15V paletas"
    MSF.TextMatrix(11, 3) = "14.0"
    MSF.TextMatrix(11, 4) = "< VM >"
    MSF.TextMatrix(11, 5) = "17.0"
    MSF.TextMatrix(11, 7) = "V"
    
    MSF.TextMatrix(12, 1) = "12"
    MSF.TextMatrix(12, 2) = "Conversión A/D tensión de batería"
    MSF.TextMatrix(12, 3) = ""
    MSF.TextMatrix(12, 4) = "VM >"
    MSF.TextMatrix(12, 5) = "14.5"
    MSF.TextMatrix(12, 7) = "V"
    
    MSF.TextMatrix(13, 1) = "13"
    MSF.TextMatrix(13, 2) = "Apagado por baja tensión"
    MSF.TextMatrix(13, 3) = "7.0"
    MSF.TextMatrix(13, 4) = "< VM >"
    MSF.TextMatrix(13, 5) = "8.0"
    MSF.TextMatrix(13, 7) = "V"
    
    MSF.TextMatrix(14, 1) = "14"
    MSF.TextMatrix(14, 2) = "Corriente de consumo apagado"
    MSF.TextMatrix(14, 3) = "100"
    MSF.TextMatrix(14, 4) = "< VM >"
    MSF.TextMatrix(14, 5) = "500"
    MSF.TextMatrix(14, 7) = "µA"
    
    MSF.TextMatrix(15, 1) = "15"
    MSF.TextMatrix(15, 2) = "Corriente de consumo encendido"
    MSF.TextMatrix(15, 3) = "0.2"
    MSF.TextMatrix(15, 4) = "< VM >"
    MSF.TextMatrix(15, 5) = "0.5"
    MSF.TextMatrix(15, 7) = "A"
    
    MSF.TextMatrix(16, 1) = "16"
    MSF.TextMatrix(16, 2) = "Corriente de carga de batería"
    MSF.TextMatrix(16, 3) = "0.1"
    MSF.TextMatrix(16, 4) = "< VM >"
    MSF.TextMatrix(16, 5) = "3.5"
    MSF.TextMatrix(16, 7) = "A"
    
    MSF.TextMatrix(17, 1) = "17"
    MSF.TextMatrix(17, 2) = "Calibración RRMC"
    MSF.TextMatrix(17, 3) = "100"
    MSF.TextMatrix(17, 4) = "< VM >"
    MSF.TextMatrix(17, 5) = "400"
    MSF.TextMatrix(17, 7) = "mVpp"
    
    MSF.TextMatrix(18, 1) = "18"
    MSF.TextMatrix(18, 2) = "Polarización ECG"
    MSF.TextMatrix(18, 3) = "120"
    MSF.TextMatrix(18, 4) = "< VM >"
    MSF.TextMatrix(18, 5) = "122"
    MSF.TextMatrix(18, 7) = "u"
    
    MSF.TextMatrix(19, 1) = "19"
    MSF.TextMatrix(19, 2) = "Ganancia ECG"
    MSF.TextMatrix(19, 3) = "112"
    MSF.TextMatrix(19, 4) = "< VM >"
    MSF.TextMatrix(19, 5) = "114"
    MSF.TextMatrix(19, 7) = "u"
    
    MSF.TextMatrix(20, 1) = "20"
    MSF.TextMatrix(20, 2) = "Selectora ECG"
    MSF.TextMatrix(20, 3) = ""
    MSF.TextMatrix(20, 4) = ""
    MSF.TextMatrix(20, 5) = ""
    MSF.TextMatrix(20, 7) = ""
    
    MSF.TextMatrix(21, 1) = "21"
    MSF.TextMatrix(21, 2) = "Captura QRS en alta frecuencia"
    MSF.TextMatrix(21, 3) = "285"
    MSF.TextMatrix(21, 4) = "< FC >"
    MSF.TextMatrix(21, 5) = "295"
    MSF.TextMatrix(21, 7) = "ppm"

    MSF.TextMatrix(22, 1) = "22"
    MSF.TextMatrix(22, 2) = "Detección de cable ECG"
    MSF.TextMatrix(22, 3) = ""
    MSF.TextMatrix(22, 4) = ""
    MSF.TextMatrix(22, 5) = ""
    MSF.TextMatrix(22, 7) = ""

    MSF.TextMatrix(23, 1) = "23"
    MSF.TextMatrix(23, 2) = "Funcionamiento circuito de carga desfibrilador"
    MSF.TextMatrix(23, 3) = ""
    MSF.TextMatrix(23, 4) = ""
    MSF.TextMatrix(23, 5) = ""
    MSF.TextMatrix(23, 7) = ""

    MSF.TextMatrix(24, 1) = "24"
    MSF.TextMatrix(24, 2) = "Funcionamiento relé de transferencia"
    MSF.TextMatrix(24, 3) = ""
    MSF.TextMatrix(24, 4) = ""
    MSF.TextMatrix(24, 5) = ""
    MSF.TextMatrix(24, 7) = ""

    MSF.TextMatrix(25, 1) = "25"
    MSF.TextMatrix(25, 2) = "Funcionamiento relé de descarga interna"
    MSF.TextMatrix(25, 3) = ""
    MSF.TextMatrix(25, 4) = ""
    MSF.TextMatrix(25, 5) = ""
    MSF.TextMatrix(25, 7) = ""

    MSF.TextMatrix(26, 1) = "26"
    MSF.TextMatrix(26, 2) = "Carga desfibrilador con batería externa"
    MSF.TextMatrix(26, 3) = ""
    MSF.TextMatrix(26, 4) = ""
    MSF.TextMatrix(26, 5) = ""
    MSF.TextMatrix(26, 7) = ""

    MSF.TextMatrix(27, 1) = "27"
    MSF.TextMatrix(27, 2) = "Carga desfibrilador con alimentación de red"
    MSF.TextMatrix(27, 3) = ""
    MSF.TextMatrix(27, 4) = ""
    MSF.TextMatrix(27, 5) = ""
    MSF.TextMatrix(27, 7) = ""

    MSF.TextMatrix(28, 1) = "28"
    MSF.TextMatrix(28, 2) = "Calibración de carga desfibrilador"
    MSF.TextMatrix(28, 3) = "4240"
    MSF.TextMatrix(28, 4) = "< VM >"
    MSF.TextMatrix(28, 5) = "4260"
    MSF.TextMatrix(28, 7) = "V"

    MSF.TextMatrix(29, 1) = "29"
    MSF.TextMatrix(29, 2) = "Calibración de limitador de carga"
    MSF.TextMatrix(29, 3) = "4085"
    MSF.TextMatrix(29, 4) = "VM ="
    MSF.TextMatrix(29, 5) = "4095"
    MSF.TextMatrix(29, 7) = "u"

    MSF.TextMatrix(30, 1) = "30"
    MSF.TextMatrix(30, 2) = "Máxima tensión de carga"
    MSF.TextMatrix(30, 3) = ""
    MSF.TextMatrix(30, 4) = "VM >"
    MSF.TextMatrix(30, 5) = "4370"
    MSF.TextMatrix(30, 7) = "V"

    MSF.TextMatrix(31, 1) = "31"
    MSF.TextMatrix(31, 2) = "Detección de paletas"
    MSF.TextMatrix(31, 3) = ""
    MSF.TextMatrix(31, 4) = ""
    MSF.TextMatrix(31, 5) = ""
    MSF.TextMatrix(31, 7) = ""

    MSF.TextMatrix(32, 1) = "32"
    MSF.TextMatrix(32, 2) = "Teclado paletas"
    MSF.TextMatrix(32, 3) = ""
    MSF.TextMatrix(32, 4) = ""
    MSF.TextMatrix(32, 5) = ""
    MSF.TextMatrix(32, 7) = ""

    MSF.TextMatrix(33, 1) = "33"
    MSF.TextMatrix(33, 2) = "Indicadores luminosos paletas"
    MSF.TextMatrix(33, 3) = ""
    MSF.TextMatrix(33, 4) = ""
    MSF.TextMatrix(33, 5) = ""
    MSF.TextMatrix(33, 7) = ""

    MSF.TextMatrix(34, 1) = "34"
    MSF.TextMatrix(34, 2) = "Apagado por software"
    MSF.TextMatrix(34, 3) = ""
    MSF.TextMatrix(34, 4) = ""
    MSF.TextMatrix(34, 5) = ""
    MSF.TextMatrix(34, 7) = ""

    MSF.TextMatrix(35, 1) = "35"
    MSF.TextMatrix(35, 2) = "Umbral QRS"
    MSF.TextMatrix(35, 3) = "40"
    MSF.TextMatrix(35, 4) = "< VM >"
    MSF.TextMatrix(35, 5) = "50"
    MSF.TextMatrix(35, 7) = "mV"
    
    MSF.TextMatrix(36, 1) = "36"
    MSF.TextMatrix(36, 2) = "Precalibración de energía entregada"
    MSF.TextMatrix(36, 3) = "197.0"
    MSF.TextMatrix(36, 4) = "< VM >"
    MSF.TextMatrix(36, 5) = "203.0"
    MSF.TextMatrix(36, 7) = "J"
    
    MSF.TextMatrix(37, 1) = "37"
    MSF.TextMatrix(37, 2) = "Calibración de energía entregada"
    MSF.TextMatrix(37, 3) = "357.0"
    MSF.TextMatrix(37, 4) = "< VM >"
    MSF.TextMatrix(37, 5) = "363.0"
    MSF.TextMatrix(37, 7) = "J"
    
    MSF.TextMatrix(38, 1) = "38"
    MSF.TextMatrix(38, 2) = "Tiempo de carga a 360 Joules"
    MSF.TextMatrix(38, 3) = "4.9"
    MSF.TextMatrix(38, 4) = "< VM >"
    MSF.TextMatrix(38, 5) = "5.5"
    MSF.TextMatrix(38, 7) = "s"
    
    MSF.TextMatrix(39, 1) = "39"
    MSF.TextMatrix(39, 2) = "Pérdida de energía almacenada"
    MSF.TextMatrix(39, 3) = ""
    MSF.TextMatrix(39, 4) = " VM >"
'    MSF.TextMatrix(39, 5) = "324"
    MSF.TextMatrix(39, 5) = "306"
    MSF.TextMatrix(39, 7) = "j"
    
    MSF.TextMatrix(40, 1) = "40"
    MSF.TextMatrix(40, 2) = "Demora en la entrega de pulso"
    MSF.TextMatrix(40, 3) = ""
    MSF.TextMatrix(40, 4) = "VM <"
    MSF.TextMatrix(40, 5) = "28.0"
    MSF.TextMatrix(40, 7) = "u"
    
    MSF.TextMatrix(41, 1) = "41"
    MSF.TextMatrix(41, 2) = "Calibración de medidor de energía entregada"
    MSF.TextMatrix(41, 3) = ""
    MSF.TextMatrix(41, 4) = "| VM - ENDE | <="
                             
    MSF.TextMatrix(41, 5) = "1"
    MSF.TextMatrix(41, 7) = "J"
    
    MSF.TextMatrix(42, 1) = "42"
    MSF.TextMatrix(42, 2) = "Marcado presets"
    MSF.TextMatrix(42, 3) = ""
    MSF.TextMatrix(42, 4) = ""
    MSF.TextMatrix(42, 5) = ""
    MSF.TextMatrix(42, 7) = ""
    
    MSF.TextMatrix(43, 1) = "43"
    MSF.TextMatrix(43, 2) = "Circuito de protección de ECG desde paletas"
    MSF.TextMatrix(43, 3) = ""
    MSF.TextMatrix(43, 4) = ""
    MSF.TextMatrix(43, 5) = ""
    MSF.TextMatrix(43, 7) = ""
    
    MSF.TextMatrix(44, 1) = "44"
    MSF.TextMatrix(44, 2) = "ECG desde paletas"
    MSF.TextMatrix(44, 3) = ""
    MSF.TextMatrix(44, 4) = ""
    MSF.TextMatrix(44, 5) = ""
    MSF.TextMatrix(44, 7) = ""
End Sub




Function VALIDARRESULTADO$(MSF As msFlexGrid, REG&, col&)
   COLUM% = col&
   If UCase$(MSF.TextMatrix(REG&, col&)) = "OK" Then
        VALIDARRESULTADO$ = "OK"
   
   ElseIf Abs(XVALO(MSF.TextMatrix(REG&, col&))) > 0 Then
        VALOR_a_VALIDAR = ROUND(XVALO(MSF.TextMatrix(REG&, col&)), 2)
        VALOR1 = ROUND(XVALO(MSF.TextMatrix(REG&, 3)), 2)
        VALOR2 = ROUND(XVALO(MSF.TextMatrix(REG&, 5)), 2)
        SIMBOLEFT$ = Left(MSF.TextMatrix(REG&, 4), 1)
        SIMBORIGHT$ = Right(MSF.TextMatrix(REG&, 4), 1)
        If SIMBOLEFT$ = "<" And SIMBORIGHT$ = ">" Then
          If VALOR_a_VALIDAR >= VALOR1 And VALOR_a_VALIDAR <= VALOR2 Then
           VALIDARRESULTADO$ = "OK"
           Call COLOREAR_TEXTO_SOLO_UNA_AFILA(MSF, vbBlack, REG&, COLUM%)
           Else: GoTo 50
          End If
        ElseIf SIMBOLEFT$ = ">" Then
          If VALOR_a_VALIDAR >= VALOR1 Then
           VALIDARRESULTADO$ = "OK"
           Call COLOREAR_TEXTO_SOLO_UNA_AFILA(MSF, vbBlack, REG&, COLUM%)
           Else: GoTo 50
          End If
        ElseIf SIMBORIGHT$ = "<" Then
          If VALOR_a_VALIDAR <= VALOR2 Then
           VALIDARRESULTADO$ = "OK"
           Call COLOREAR_TEXTO_SOLO_UNA_AFILA(MSF, vbBlack, REG&, COLUM%)
           Else: GoTo 50
          End If
        ElseIf SIMBORIGHT$ = ">" Then
          If SIMBOLEFT$ <> "<" Then 'AGREGADO PARA QUE SI SOLO TIENE VALOR > DEL LADO DERECHO Y NADA DEL OTRO VALIDE QUE SEA MAYOR
             If VALOR_a_VALIDAR > VALOR2 Then
             VALIDARRESULTADO$ = "OK"
             Call COLOREAR_TEXTO_SOLO_UNA_AFILA(MSF, vbBlack, REG&, COLUM%)
             Exit Function
          Else
              GoTo 50
          End If
           
          If VALOR_a_VALIDAR >= VALOR2 Then
             VALIDARRESULTADO$ = "OK"
             Call COLOREAR_TEXTO_SOLO_UNA_AFILA(MSF, vbBlack, REG&, COLUM%)
             Else: GoTo 50
            
           End If
          End If
        Else
50         VALIDARRESULTADO$ = "MAL"
           Call COLOREAR_TEXTO_SOLO_UNA_AFILA(MSF, vbRed, REG&, COLUM%)
        End If
  End If
End Function

Private Sub TXTMOVIL_LostFocus()
'     If TXTMOVIL.Visible = True Then
'       REG& = XVALO(TXTMOVIL.ToolTipText)
''       MSF1.TextMatrix(REG&, 6) = FORMATNUM$(XVALO(TXTMOVIL), "F12.2")
''       TXTMOVIL.Visible = False
'     End If
'
'      IDFILA& = XVALO(MSF.TextMatrix(REG&, 2))
'
'      If (IDFILA& >= 6 And IDFILA& <= 17) Or (IDFILA& = 23) Or (IDFILA& >= 37 And IDFILA& <= 42) Then
'         MSF.TEXT = TRIM$(FORMATNUM$(XVALO(TXTMOVIL.TEXT), "F12.2"))
'         REG& = XVALO(TXTMOVIL.ToolTipText)
'         MSF.TextMatrix(REG&, 0) = VALIDARRESULTADO$(MSF, REG&, 6)
'''
'      Else
'         MSF.TEXT = TRIM$(TXTMOVIL.TEXT)
'         REG& = XVALO(TXTMOVIL.ToolTipText)
'         MSF.TextMatrix(REG&, 0) = VALIDARRESULTADO$(MSF, REG&, 6)
'      End If

End Sub
