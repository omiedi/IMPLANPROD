VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form RegCFinEym 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "REGISTRO DE CONTROL FINAL"
   ClientHeight    =   9195
   ClientLeft      =   45
   ClientTop       =   915
   ClientWidth     =   14250
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9195
   ScaleWidth      =   14250
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "Rep"
      Height          =   300
      Left            =   9450
      TabIndex        =   42
      Top             =   315
      Width           =   520
   End
   Begin VB.TextBox TXTCI 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   8295
      Locked          =   -1  'True
      TabIndex        =   38
      Text            =   " "
      Top             =   0
      Width           =   960
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   10680
      Locked          =   -1  'True
      TabIndex        =   31
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
      Left            =   3840
      TabIndex        =   30
      TabStop         =   0   'False
      ToolTipText     =   "Costo segun el proveedor y la via de entrega"
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Height          =   750
      Left            =   3960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   990
      Width           =   6570
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      Height          =   300
      Left            =   9480
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1320
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox Text11 
      Appearance      =   0  'Flat
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
      TabIndex        =   12
      Top             =   0
      Width           =   930
      Begin VB.CommandButton Command23 
         Height          =   650
         Left            =   105
         Picture         =   "RegCFinEym.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Visualiza e Imprime Informe de  Fuga a Tierra"
         Top             =   2280
         Width           =   650
      End
      Begin VB.CommandButton Command26 
         Height          =   650
         Left            =   105
         Picture         =   "RegCFinEym.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Configuración de Formulario"
         Top             =   1560
         Width           =   650
      End
      Begin VB.CommandButton Command2 
         Height          =   650
         Left            =   120
         Picture         =   "RegCFinEym.frx":0AAC
         Style           =   1  'Graphical
         TabIndex        =   7
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
         TabIndex        =   13
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
            TabIndex        =   14
            Top             =   120
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command29 
         Height          =   630
         Left            =   105
         Picture         =   "RegCFinEym.frx":0DB6
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Ayuda Flexoft en Línea"
         Top             =   840
         Width           =   650
      End
      Begin VB.CommandButton Command1 
         Height          =   640
         Left            =   105
         Picture         =   "RegCFinEym.frx":10C0
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Salir"
         Top             =   120
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "RegCFinEym.frx":120A
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   7970
         Width           =   1515
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Bindings        =   "RegCFinEym.frx":312C
      Height          =   7365
      Left            =   120
      TabIndex        =   28
      Top             =   1800
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
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Asist.Técnica"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   7440
      TabIndex        =   43
      Top             =   650
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Label LBLREVISION 
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
      Left            =   6375
      TabIndex        =   41
      Top             =   675
      Width           =   1185
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Rev.:"
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
      Index           =   7
      Left            =   5880
      TabIndex        =   40
      Top             =   765
      Width           =   405
   End
   Begin VB.Label Label14 
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
      TabIndex        =   39
      Top             =   45
      Width           =   1260
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "'NC' (No Conforme)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      TabIndex        =   37
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   15.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   1800
      TabIndex        =   36
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   15.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   1800
      TabIndex        =   35
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "'C' (Conforme)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      TabIndex        =   34
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   495
      Index           =   0
      Left            =   6480
      TabIndex        =   33
      Top             =   4320
      Width           =   1215
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
      TabIndex        =   32
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
      Left            =   0
      TabIndex        =   29
      Top             =   720
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   10245
      TabIndex        =   27
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
      Left            =   9900
      TabIndex        =   26
      Top             =   420
      Width           =   300
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Realizó"
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
      TabIndex        =   25
      Top             =   520
      Width           =   1380
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   10680
      TabIndex        =   24
      ToolTipText     =   "Doble Click Para Seleccionar Personal de Aprobación"
      Top             =   1320
      Width           =   2520
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   120
      Picture         =   "RegCFinEym.frx":3140
      Top             =   1080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "RegCFinEym.frx":344A
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
      TabIndex        =   23
      Top             =   360
      Width           =   810
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Veredicto General:"
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
      Left            =   9600
      TabIndex        =   22
      Top             =   810
      Visible         =   0   'False
      Width           =   900
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
      TabIndex        =   21
      Top             =   360
      Width           =   1260
   End
   Begin VB.Label Label95 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   960
      TabIndex        =   20
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
      TabIndex        =   19
      Top             =   120
      Width           =   780
   End
   Begin VB.Label Label88 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   240
      TabIndex        =   18
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
      TabIndex        =   17
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
      TabIndex        =   16
      Top             =   720
      Width           =   1530
   End
   Begin VB.Label Label4 
      BackColor       =   &H00CDDADA&
      Caption         =   "REGISTRO DE CONTROL FINAL"
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
      TabIndex        =   15
      Top             =   0
      Width           =   3255
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
Attribute VB_Name = "RegCFinEym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TXTRAZA_COFUTIE$
Dim isMaximized As Boolean

Sub Actualizar_Validaciones_En_Grilla()

   For i& = 1 To MSF.Rows - 1
       REG& = i&
       IDFILA& = MSF.TextMatrix(REG&, 1)
       If (IDFILA& >= 10 And IDFILA& <= 15) Or IDFILA& = 17 Or IDFILA& = 34 Then
        COLX& = 3: COLUM% = 3
        MSF.TextMatrix(REG&, 8) = VALIDARRESULTADO$(MSF, REG&, COLX&)
      Else
        COLX& = 8: COLUM% = 8
        'MSF.TextMatrix(REG&, 8) = VALIDARRESULTADOOKYMAL$(MSF, REG&, COLX&)
        If MSF.TextMatrix(REG&, 8) = "MAL" Then
           Call COLOREAR_TEXTO_SOLO_UNA_AFILA(MSF, vbRed, REG&, COLUM%)
        ElseIf MSF.TextMatrix(REG&, 8) = "OK" Then
           Call COLOREAR_TEXTO_SOLO_UNA_AFILA(MSF, vbBlack, REG&, COLUM%)
        End If
      End If

   Next i&

End Sub

Sub ActualizarOperacion34(NUSERIE$)
   CONDI$ = "(OP = 0 OR OP IS NULL)  and numeroserie = '" & NUSERIE$ & "' AND (DESCRIPCION = 'Demora en modo sincronizado' OR DESCRIPCION = 'Presencia de espiga en QRS')"
   
   CANTI = CantRegistros("CONTFINEYM", CONDI$, "NOMESS")
   If CANTI >= 1 And CANTI <= 10 Then ' menor a 10 por las dudas.
      Call ACTUREGISTRO("CONTFINEYM", "OP", CONDI$, 34, REGAF&, "NOMESS")
   End If
   
End Sub

Function SiAplicaoNoEym$(NroFilaEym%, Modelo$)
  'EJEMPLO VA  A VENIR EN MODELO$ = DB0001
  'QUEDO SIN USO SE ESTA CARGANDO EN LA TABLA CONTROEYM
  'SE CARGA LA CONFIGURACION DESDE FORMULARIO LOTEXSQR > TRANSACCIONES > ENGANCHARE N.SERIE POR ORDEN DE FABRICACION > mnuCONFIGURACION > CONTROLES POR EQUIPOS
  ModeloNum% = XVALO(Mid$(Modelo$, 3))
  NumModelo$ = TRIM$(CStr(ModeloNum%))
  NombreCampoDataSql$ = "Campo" & NumModelo$
  NumRegistroDataSql$ = TRIM$(CStr(NroFilaEym%))
  
   'consulta a la tabla datasql por  el numero de fila que esta en el campo20 (Nombre: NroRegistro)
   'y segun el nro de modelo que va del DB0001 al DB0016 busca la columna con dicha numeracion para el DB0001 = campo1
   'le devuelve si aplica o no.
   'si la select no devuelve ningun registro va como N/A
   SQLX$ = "SELECT " & NombreCampoDataSql$ & " From DataSql"
   SQLX$ = SQLX$ + " where Campo20= '" & NumRegistroDataSql$ & "'"
   SQLX$ = SQLX$ + " "
   CONDI$ = "LISTA = 'CONFIGURACION C.F. C15' AND Campo20= '" & NumRegistroDataSql$ & "'"
   SiAplicaoNoEym$ = TRIM$(VALOBADA("DATASQL", NombreCampoDataSql$, CONDI$, "NOMESS"))
   If SiAplicaoNoEym$ = "" Then SiAplicaoNoEym$ = "N/A"
End Function

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
    If TRIM$(Text1) = "" Then
      Call MENSERR(24, "Falta Informar Quién Controló")
      GoTo 99
    End If

    Screen.MousePointer = 11
    '

    ULTREVIS% = XVALO(VALOBADA("CONTFINEYM", "MAX(REVISION)", "NumeroSerie ='" & NUSERIE$ & "'"))
    REVIS% = ULTREVIS% + 1

    CIXI% = CODINT%(Label95)
    RESULTADOX$ = "C"
    EXCEP$ = EXCEPCION_CONTROL_EYM$(CIXI%, "CONTFINEYM")
    Call TEXTOLOTES(EXCEP$, ";")
    '
    'ACA VALIDO LOS DATOS LEIDOS CONSIDERANDO LAS EXCEPCIONES CONFIGURADAS
    'SALVO EXCEPCIONES DEBE ESTAR TODO OK PARA QUE ESTE CONFORME
    'CASO CONTRARIO EL RESULDATO ES NC
    
    
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
           If TRIM$(UCase$(MSF.TextMatrix(i&, 8))) <> "N/A" Then
                Label5.Caption = Chr(254)
                Label12.Caption = Chr(168)
                Label5.Refresh
                Label12.Refresh
                RESULTADOX$ = "NC"
                Exit For
           End If
        End If
20  Next i&
    Call BLOQUEAR_CONTROLES(Me, 2, 1, 1, , 1, 1)
    On Error GoTo ERROR_GUARDAR
    FLEXCONN.BeginTrans

    SQLX$ = "SELECT * FROM CONTFINEYM"
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
            !NUMEROSERIE = TRIM$(Text8)
            !Modelo = CIXI%
            !COD_EXTE = TRIM$(Me.Label95)
            !OP = XVALO(MSF.TextMatrix(i&, 1))
            !Descripcion = MSF.TextMatrix(i&, 2)
            !Rango_Acep_Min = MSF.TextMatrix(i&, 5)
            !VM = MSF.TextMatrix(i&, 6)
            !Rango_Acep_Max = MSF.TextMatrix(i&, 7)
            !OK_O_MAL = MSF.TextMatrix(i&, 8)
            !VALOR_RESUL = MSF.TextMatrix(i&, 3)
            !Unimed = MSF.TextMatrix(i&, 4)
            !Bloqueado = "SI"
            !Revision = REVIS%
            !Controlo = Left$(Text1, 32)
            !Comentarios = Left$(Text9, 512)
            !APROBO = Left$(Label8, 32)
            !IDSUBOR = Left$(Text2, 16)
            !FECH_ENSA = CVDAT(FECHANUM(Text11))
            !Resultado = RESULTADOX$
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
    
    Call BLOQUEAR_CONTROLES(Me, 1, 1, 1, , 1, 1)

    Call PRINTREGCF(NUSERIE$, REVIS%)
    '
    
    Image1.Visible = True
    Image3.Visible = False
    Command23.Enabled = True
    '
    Call BLANFORMU
    '
    Call CARGA_REGISTRO_ContFinal
    Screen.MousePointer = 1
    '
99  Command2.Enabled = True
   '\\\OT-08-0936-OD-fin///
End Sub
Sub PRINTREGCF(NUSERIE$, ULTREVIS%)
    FORIMPRE$ = CONTROL$("", "CTFINEYM")

    
    If FORIMPRE$ = "" Then
       Call COMUNI("Falta Configurar Formulario de Impresión de Control Final")
       Exit Sub
    End If
    '
    NORFA$ = TRIM$(Text2)
    If NORFA$ = "" Then
      OPX1% = COMALTER%("No Hay una Orden de Fabricación Activa\\Cancelar\Continuar")
      If OPX1% < 2 Then Exit Sub
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
    PDES$ = CONTROL$(IDENTER$, "PRIREGCF") 'la configuracion ya esta.
    If PDES$ = "" Then PDES$ = CONTROL$("*", "PRIREGCF")
    
    If Label15.Visible = True Then 'ES REPARACION
      RUTAEXCEL$ = CONTROL$("REPARACI", "CTFIREPA")
    Else
      RUTAEXCEL$ = CONTROL("REGICFIN", "ARCHCFIN")
    End If
    If EXISTE%(RUTAEXCEL$) < 1 Then
       Call COMUNI("Falta Configurar Ruta de Archivo Excel de Control Final")
       Exit Sub
    End If
    
    With VENTABNEW.MSF
    For i& = 1 To Me.MSF.Rows - 1
          .Rows = i& + 1
        .TextMatrix(i&, 1) = "" 'LIMPIO ESTAS DOS FILAS P.Q. NO LAS LEO Y SEGUN EL CASO LAS ESCRIBO
         .TextMatrix(i&, 2) = ""
       Select Case i&
         Case 10, 11, 12, 13, 14, 15, 17, 34
            .TextMatrix(i&, 2) = Me.MSF.TextMatrix(i&, 3)
         Case Else
            If Me.MSF.TextMatrix(i&, 8) = "OK" Then
               .TextMatrix(i&, 1) = "A"
            ElseIf Me.MSF.TextMatrix(i&, 8) = "MAL" Then
               .TextMatrix(i&, 1) = "M"
            ElseIf Me.MSF.TextMatrix(i&, 8) = "N/A" Then
               .TextMatrix(i&, 1) = "A"
               If i& = 35 Then
                  .TextMatrix(i&, 1) = "N/A"
               End If
            Else
               .TextMatrix(i&, 1) = ""
            End If
       End Select
    Next i&
    End With
'   k& = MSF.Rows
'    VENTABNEW.Show 1
    REVISIONX$ = "-"
    If NORFA$ = "" Then
       REVISIONX$ = LBLREVISION.Caption
    End If
    
    For i& = 1 To 7
      Select Case i&
         Case 1: VENTABNEW.MSF.TextMatrix(i, 3) = Label95  'CODIGO
         Case 2: VENTABNEW.MSF.TextMatrix(i, 3) = LBLREVISION.Caption   'REVISION
         Case 3: VENTABNEW.MSF.TextMatrix(i, 3) = NSERIE$   'N.SERIE
         Case 4: VENTABNEW.MSF.TextMatrix(i, 3) = Text11   'FECHA
         Case 5: VENTABNEW.MSF.TextMatrix(i, 3) = Text1  'REALIZO
         Case 6: VENTABNEW.MSF.TextMatrix(i, 3) = Text9  'COMENTARIOS
      End Select
    Next i&
    '
        
        
    Call Param_Print_Excel(RUTAEXCEL$, "111", VENTABNEW.MSF, PDES$, "1", 12, 1, "", 1, 1)     'imprime en excel la primera columna de mi msf en la columna 1 a partir de la fila 12 en el excel (este graba pero no lo imprime)
    Call Param_Print_Excel(RUTAEXCEL$, "111", VENTABNEW.MSF, PDES$, "2", 12, 7, "", 1, 1) 'imprime en excel la 2 columna de mi msf en la columna 8 a partir de la fila 12 en el excel (este graba pero no lo imprime)
    Call Param_Print_Excel(RUTAEXCEL$, "111", VENTABNEW.MSF, PDES$, "3", 54, 22, "", 1, 0)   'imprime en excel la 3 columna de mi msf en la columna 1 a partir de la fila 54 en el excel  (este imprime con los 3 cambios )
    Unload VENTABNEW
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
    Call PRINTREGCF(NUSERIE$, ULTREVIS%)
    '
End Sub
Private Sub Command26_Click()
   Command26.Enabled = False
   FSTPRCF_EYM.Show 1
   Command26.Enabled = True
End Sub

Private Sub Command3_Click()
   On Error Resume Next
   RESP_ZELE_VECT(2) = ""
   On Error GoTo 0
   Campos$ = "Reparación,N°Serie/A12;Código/A16;Descripción/A28;F.Alta/D8"
   Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO)
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
   FRMZELECHO.Caption = ""
'  WHERE NUMREPA = " & NREP&
   SQLX$ = "SELECT R.NUMREPA, R.NUMEROSERIE,M.CODIGO,R.DESCRIPCION,R.FECHENT FROM REPARA R WITH(NOLOCK)"
   SQLX$ = SQLX$ + " INNER JOIN MASTER M ON M.CODINT = R.COD_INTE WHERE R.STATUS >= 3"
   SQLX$ = SQLX$ + " ORDER BY R.NUMREPA desc"
   Set rs1 = READSET(SQLX$)
   Call Carga_MSF_Recordset(rs1, Campos$, FRMZELECHO.msf2)
   Set obj = Nothing
   rs1.Close
   Set rs1 = Nothing
   Label15.Visible = False

   FRMZELECHO.Width = 12500
   FRMZELECHO.Show 1
   If TRIM$(RESP_ZELE_VECT(2)) <> "" Then
      NSERIE$ = TRIM$(RESP_ZELE_VECT(2))
      Text8 = NSERIE$
      Label95 = TRIM$(RESP_ZELE_VECT(3))
      Label88 = TRIM$(RESP_ZELE_VECT(4))
      TXTCI = CODINT%(Label95)
      Label15.Visible = True
      Call Text8_KeyPress(13)
      
   End If

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
    Call CARGA_REGISTRO_ContFinal
    Label12 = Chr(168)
    Label5 = Chr(168)
    ' Establece el estado inicial de la variable isMaximized
    isMaximized = True

    '\\\OT-08-0936-OD-FIN///
End Sub

Private Sub Form_Resize()

    If isMaximized Then
        ' Si el formulario ya está maximizado, restaura su tamaño original
        Me.WindowState = vbNormal
        MSF.Height = 7365
        MSF.Width = 13080
        Picture13.Left = 13365
        isMaximized = False
    Else
        ' Si el formulario no está maximizado, ajusta su tamaño para ocupar toda la pantalla
        Me.WindowState = vbMaximized
        MSF.Height = Me.Height - (Me.Height * 0.265)
        MSF.Width = Me.Width - (Me.Picture13.Width * 1.2)
        Picture13.Left = Me.Width - Me.Picture13.Width
        Picture13.Height = Me.Height - (Me.Height * 0.1)
        isMaximized = True
    End If
End Sub

Private Sub Frame5_DragDrop(Source As CONTROL, X As Single, Y As Single)
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
          ULTIMOENSA% = XVALO(VALOBADA("CONTFINEYM ", "MAX(REVISION)", CONDI$))
          CONDI$ = CONDI$ + " AND REVISION = " & ULTIMOENSA%
          Text11 = FECHATEX(CVINT(VALOBADA("CONTFINEYM", "Fech_Ensa", CONDI$, "NOMESS")))
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
        Text11 = FECHATEX(HOII%)
       
   End If
End Sub

Private Sub Label12_Click()
    If Image3.Visible = True Then Exit Sub ' SI ESTA BLOQUEADO NO CAMBIA
    NUSERIE$ = TRIM$(Text8)
    If Label12.Caption = Chr(254) Then
       Label12.Caption = Chr(168)
    Else
       Label12.Caption = Chr(254)
       Label5.Caption = Chr(168)
    End If

   '
End Sub



Private Sub Label5_Click()
    If Image3.Visible = True Then Exit Sub ' SI ESTA BLOQUEADO NO CAMBIA
    NUSERIE$ = TRIM$(Text8)
    If Label6.Caption = Chr(254) Then
       Label5.Caption = Chr(168)
    Else
       Label5.Caption = Chr(254)
       Label12.Caption = Chr(168)
    End If

End Sub


Private Sub Label8_Click()
    If Image3.Visible = True Then Exit Sub ' SI ESTA BLOQUEADO NO CAMBIA
    If Image3.Visible = True Then Exit Sub ' SI ESTA BLOQUEADO NO CAMBIA
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
        EXCEP$ = EXCEPCION_CONTROL_EYM$(CIXI%, "CONTFINEYM")
        If TRIM$(EXCEP$) = "" Then Exit Sub
        Call TEXTOLOTES(EXCEP$, ";")
        '
        For i& = 1 To MSF.Rows - 1
            On Error Resume Next
            'VALIDA SI TIENE EXCEPCIONES
            J& = 0
            While J& < UBound(CADENATEX$)
              J& = J& + 1
              If XVALO(MSF.TextMatrix(i&, 1)) = XVALO(CADENATEX$(J&)) Then
                   MSF.TextMatrix(i&, 8) = "N/A"
                   Call COLOREAR_GRILLA_SOLO_UNA_AFILA(MSF, &HCDDADA, i&)
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
    
    If MSF.TextMatrix(REG&, 8) = "N/A" And REG& <> 35 And Label15.Visible = False Then Exit Sub

    If COLUM% <> 3 And COLUM% <> 8 Then
      TXTMOVIL.Visible = False
      Exit Sub
    End If
    '
    
    IDFILA& = REG&
    'If (IDFILA& >= 3 And IDFILA& <= 17) Or (IDFILA& = 23) Or (IDFILA& >= 37 And IDFILA& <= 42) Then
    If (IDFILA& >= 10 And IDFILA& <= 15) Or IDFILA& = 17 Or IDFILA& = 34 Then
        MSF.COL = 3
        With TXTMOVIL
          .ZOrder 0
          .Width = MSF.CellWidth
          .Left = MSF.CellLeft + MSF.Left
          .Top = MSF.CellTop + MSF.Top
          .Height = MSF.CellHeight
          .Visible = True
          .ToolTipText = MSF.Row
          .TEXT = MSF.TEXT
          .SetFocus
        End With
        COLX& = 3:: COLUM% = 3
        MSF.TextMatrix(REG&, 8) = VALIDARRESULTADO$(MSF, REG&, COLX&)
    Else
        COLX& = 8: COLUM% = 8
        MSF.TextMatrix(REG&, 8) = VALIDARRESULTADOOKYMAL$(MSF, REG&, COLX&)
        If MSF.TextMatrix(REG&, 8) = "MAL" Then
           Call COLOREAR_TEXTO_SOLO_UNA_AFILA(MSF, vbRed, REG&, COLUM%)
        Else
           Call COLOREAR_TEXTO_SOLO_UNA_AFILA(MSF, vbBlack, REG&, COLUM%)
        End If
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

Private Sub TXTCI_Change()
    LBLREVISION.Caption = REVICOD$(XVALO(Me.TXTCI))
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
    COLU1& = 3
    MSF.COL = COLU1&
    'RUTINA PARA QUE AL DAR ENTRER O ESCAPE SAQUE EL FOCO Y PONGA COMO NO VISIBLE EL TEXT DE EDICION
    If KeyAscii = 13 Then
      REG& = XVALO(TXTMOVIL.ToolTipText)
      MSF.Row = REG&
'      If TRIM$(MSF.TextMatrix(REG&, 3)) = "" And TRIM$(MSF.TextMatrix(REG&, 3)) = "" Then
      'IDFILA& = XVALO(MSF.TextMatrix(REG&, 1))
         IDFILA& = REG&
      If (IDFILA& >= 10 And IDFILA& <= 15) Or IDFILA& = 17 Or IDFILA& = 34 Then
         MSF.TEXT = TRIM$(FORMATNUM$(XVALO(TXTMOVIL.TEXT), "F12.2"))
         REG& = XVALO(TXTMOVIL.ToolTipText)
         MSF.TextMatrix(REG&, 8) = VALIDARRESULTADO$(MSF, REG&, 3)
      Else
         GoTo 20
         MSF.TEXT = TRIM$(TXTMOVIL.TEXT)
         REG& = XVALO(TXTMOVIL.ToolTipText)
         MSF.TextMatrix(REG&, 8) = VALIDARRESULTADO$(MSF, REG&, 3)
      End If
20    TXTMOVIL.Visible = False
      On Error Resume Next
'      REG& = XVALO(CADENATEX$(1))
      On Error GoTo 0
      If REG& >= MSF.Rows - 1 Then
        Exit Sub
      Else
         MSF.Row = REG& + 1
         MSF.COL = 3
         With TXTMOVIL
          .ToolTipText = MSF.Row
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



Private Sub Text3_Change(Index As Integer)
   
End Sub
Private Sub Text3_GotFocus(Index As Integer)
   If Image3.Visible = True Then Command4.SetFocus
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
   Command23.Enabled = True
   TXTRAZA_COFUTIE$ = ""
   TXTCI = ""
   Call LIMPIAINFORMACION
   
   '\\\OT-08-0936-OD-08/10/08///
End Sub
Private Sub Text8_DblClick()
   '\\\OT-08-0936-OD-08/10/08///
    NSERIE$ = SERIES_CONTROLFINAL
    If NSERIE$ <> "" Then
      Text8 = NSERIE$
      Label95 = TRIM$(RESP_ZELE_VECT(2))
      Label88 = TRIM$(RESP_ZELE_VECT(3))
      Text2 = TRIM$(RESP_ZELE_VECT(4))
      Text11 = TRIM$(RESP_ZELE_VECT(5))
      TXTCI = CODINT%(Label95)
      Label15.Visible = False

      Call Text8_KeyPress(13)
      'Call Actualizar_Validaciones_En_Grilla
    End If
 
   '\\\OT-08-0936-OD-08/10/08///
End Sub
 Function SERIES_CONTROLFINAL()

   On Error Resume Next
   RESP_ZELE_VECT(1) = ""
   On Error GoTo 0
   Campos$ = "N°Serie/A12;Código/A16;Descripción/A28;O.Fabric./A12;F.Alta/D8"
   Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO)
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
   FRMZELECHO.Caption = ""
 
   Dim obj As New RECORXET
   Set rs1 = obj.RS_COD_LLEVAN_CONTROL("CONTFINEYM")
   Call Carga_MSF_Recordset(rs1, Campos$, FRMZELECHO.msf2)
   Set obj = Nothing
   rs1.Close
   Set rs1 = Nothing

   FRMZELECHO.Width = 10500
   FRMZELECHO.Show 1
   If TRIM$(RESP_ZELE_VECT(1)) <> "" Then
      SERIES_CONTROLFINAL = TRIM$(RESP_ZELE_VECT(1))
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
    Label12 = "242 V"
'    Label98 = "I354"
'    RIGIDIEL07.Label98 = ""
'    RIGIDIEL07.Text5 = ""
    
End Sub
Sub LIMPIAINFORMACION()
    Text1 = "" '\\\OT-08-1065-OD-21/11/08-MEIDTEA-LIMPIA EL TEXT DEL TECNIO
    Text7 = ""
    Text9 = ""
    Label13 = "" '\\\OT-08-1064-OD-21/11/08-MEDITEA-LIMPIO EL LABEL DE REVISION
    Label88 = ""
    Label95 = ""
    Label98 = ""
    Label8 = ""
    Label15.Visible = False
    Call LIMPIAR_MSF_1_COLUMNA_SIN_BORRAR_REGISTROS(MSF, 8)
    Call VOLVER_ABLANCO_TODA_LA_GRILLA(MSF)

End Sub

Private Sub Text8_GotFocus()
'  Call Command4.SetFocus
'  Call Command4_Click
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
   
             NUSERIE$ = TRIM$(Text8)
             Call ActualizarOperacion34(NUSERIE$)
             
'             CANTRE& = CantRegistros&("TANUMSE", "NumeroSerie ='" & NUSERIE$ & "'")
'             If CANTRE& < 1 Then
'               Call MENSERR(24, "Número de Serie Incorrecto")
'               Exit Sub
'             End If
'             '
             Screen.MousePointer = 11
             
             ULTREVIS% = XVALO(VALOBADA("CONTFINEYM", "MAX(REVISION)", "NumeroSerie ='" & NUSERIE$ & "'AND MODELO = " & XVALO(TXTCI)))
             Label13 = ULTREVIS%
            
             BLQUEO$ = "": VUELTA% = 0
             SQLX$ = "SELECT  *  FROM CONTFINEYM WITH(NOLOCK)"
             SQLX$ = SQLX$ + " WHERE NumeroSerie ='" & NUSERIE$ & "'"
             SQLX$ = SQLX$ + " AND REVISION = " & ULTREVIS% & ""
             SQLX$ = SQLX$ + " order by OP "
             Set rs = READSET(SQLX$)
             With rs
               Do While Not .EOF
                      i& = i& + 1
                      
                      MSF.TextMatrix(i&, 1) = XVALO(!OP)
                      MSF.TextMatrix(i&, 2) = SAFETEXT$(!Descripcion)
                      MSF.TextMatrix(i&, 3) = SAFETEXT$(!VALOR_RESUL)
                      MSF.TextMatrix(i&, 4) = SAFETEXT$(!Unimed)
                      MSF.TextMatrix(i&, 5) = SAFETEXT$(!Rango_Acep_Min)
                      'MSF.TextMatrix(i&, 6) = SAFETEXT$(!VM)
                      MSF.TextMatrix(i&, 7) = SAFETEXT$(!Rango_Acep_Max)
                      MSF.TextMatrix(i&, 8) = SAFETEXT$(!OK_O_MAL)
                      '
                      If VUELTA% < 1 Then
                        Text11 = FECHATEX$(CVINT(!FECH_ENSA))
                        BLQUEO$ = SAFETEXT$(!Bloqueado)
                        Text1 = SAFETEXT$(!Controlo)
                        Text9 = SAFETEXT$(!Comentarios)
                        Label8 = SAFETEXT$(!APROBO)
                        VL% = 168: VLL% = 168
                        VL1$ = TRIM$(SAFETEXT$(!Resultado))
                        If VL1$ = "C" Then
                          Label12 = Chr$(254)
                          Label5 = Chr$(168)
                        Else
                          Label12 = Chr$(168)
                          Label5 = Chr$(254)
                        End If
                        
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
        
    End If
End Sub
Sub PresentaListaSeleccionSerieEYM()
'SELECT T.COD_EXTE,T.DESCRIPCION,T.NUMEROSERIE,T.IDSUBOR FROM DATASQL D INNER JOIN TRAZAXSERIE T ON D.CAMPO1=T.COD_EXTE WHERE D.LISTA = 'CODMODELO'   GROUP BY T.COD_EXTE,T.DESCRIPCION,T.NUMEROSERIE, T.IDSUBOR
'ORDER BY NUMEROSERIE DESC

End Sub

Private Sub Text9_GotFocus()
   If Image3.Visible = True Then Command4.SetFocus
End Sub

Sub CARGA_REGISTRO_ContFinal()
    Call MSF.Clear
    Campos$ = "Item./F5;Descripción/A48;Registro/A6;U./A6;Rango/F6.2;Rango/F6.2;Rango/F6.2;Registro/F12.2;OK/A6"
    Call CARGA_ENCABEZADO(MSF, Campos$, Me, , , , 7)
    
    MSF.MergeRow(0) = True
    MSF.MergeCells = flexMergeFree
    MSF.Rows = 43
    MSF.FixedRows = 1
    MSF.Cols = 9

    MSF.TextMatrix(1, 1) = "1"
    MSF.TextMatrix(1, 2) = "Inspección visual"
    
    MSF.TextMatrix(2, 1) = "2"
    MSF.TextMatrix(2, 2) = "Trabapaletas"
    
    MSF.TextMatrix(3, 1) = "3"
    MSF.TextMatrix(3, 2) = "Alimentación desde red de c.a."
    
    MSF.TextMatrix(4, 1) = "4"
    MSF.TextMatrix(4, 2) = "Cargador de batería"
    
    MSF.TextMatrix(5, 1) = "5"
    MSF.TextMatrix(5, 2) = "Indicadores Luminosos"
    
    MSF.TextMatrix(6, 1) = "6"
    MSF.TextMatrix(6, 2) = "Tonos de audio"
    
    MSF.TextMatrix(7, 1) = "7"
    MSF.TextMatrix(7, 2) = "Configuración"
    
    MSF.TextMatrix(8, 1) = "8"
    MSF.TextMatrix(8, 2) = "Display"
    
    MSF.TextMatrix(9, 1) = "9"
    MSF.TextMatrix(9, 2) = "Test externo (x5)"
    
    MSF.TextMatrix(10, 1) = "10"
    MSF.TextMatrix(10, 2) = "Entrega 2 Joules"
    MSF.TextMatrix(10, 3) = ""
    MSF.TextMatrix(10, 4) = "J"
    MSF.TextMatrix(10, 5) = "1.5"
    MSF.TextMatrix(10, 6) = "< VM >"
    MSF.TextMatrix(10, 7) = "2.5"
    MSF.TextMatrix(10, 8) = ""
    
    MSF.TextMatrix(11, 1) = "11"
    MSF.TextMatrix(11, 2) = "Entrega 10 Joules"
    MSF.TextMatrix(11, 3) = ""
    MSF.TextMatrix(11, 4) = "J"
    MSF.TextMatrix(11, 5) = "9"
    MSF.TextMatrix(11, 6) = "< VM >"
    MSF.TextMatrix(11, 7) = "11"
    
    MSF.TextMatrix(12, 1) = "12"
    MSF.TextMatrix(12, 2) = "Entrega 50 Joules"
    MSF.TextMatrix(12, 3) = ""
    MSF.TextMatrix(12, 4) = "J"
    MSF.TextMatrix(12, 5) = "47"
    MSF.TextMatrix(12, 6) = "< VM >"
    MSF.TextMatrix(12, 7) = "53"
    
    MSF.TextMatrix(13, 1) = "13"
    MSF.TextMatrix(13, 2) = "Entrega 200 Joules"
    MSF.TextMatrix(13, 3) = ""
    MSF.TextMatrix(13, 4) = "J"
    MSF.TextMatrix(13, 5) = "190"
    MSF.TextMatrix(13, 6) = "< VM >"
    MSF.TextMatrix(13, 7) = "210"
    MSF.TextMatrix(13, 8) = ""
    
    MSF.TextMatrix(14, 1) = "14"
    MSF.TextMatrix(14, 2) = "Entrega 300 Joules"
    MSF.TextMatrix(14, 3) = ""
    MSF.TextMatrix(14, 4) = "J"
    MSF.TextMatrix(14, 5) = "285"
    MSF.TextMatrix(14, 6) = "< VM >"
    MSF.TextMatrix(14, 7) = "315"
    MSF.TextMatrix(14, 8) = ""
    
    MSF.TextMatrix(15, 1) = "15"
    MSF.TextMatrix(15, 2) = "Entrega 360 Joules"
    MSF.TextMatrix(15, 3) = ""
    MSF.TextMatrix(15, 4) = "J"
    MSF.TextMatrix(15, 5) = "342"
    MSF.TextMatrix(15, 6) = "< VM >"
    MSF.TextMatrix(15, 7) = "378"
    MSF.TextMatrix(15, 8) = ""
    
    MSF.TextMatrix(16, 1) = "16"
    MSF.TextMatrix(16, 2) = "Entrega interna 360 Joules (x10)"
    MSF.TextMatrix(16, 3) = ""
    MSF.TextMatrix(16, 4) = ""
    MSF.TextMatrix(16, 5) = ""
    MSF.TextMatrix(16, 7) = ""
    MSF.TextMatrix(16, 8) = ""
    
    MSF.TextMatrix(17, 1) = "17"
    MSF.TextMatrix(17, 2) = "Tiempo de carga a 360 Joules"
    MSF.TextMatrix(17, 3) = ""
    MSF.TextMatrix(17, 4) = "S"
    MSF.TextMatrix(17, 5) = ""
    MSF.TextMatrix(17, 6) = "VM <"
    MSF.TextMatrix(17, 7) = "6.0"
    MSF.TextMatrix(17, 8) = ""
    
    MSF.TextMatrix(18, 1) = "18"
    MSF.TextMatrix(18, 2) = "Descarga automática de energía acumulada"
    MSF.TextMatrix(18, 3) = ""
    MSF.TextMatrix(18, 4) = ""
    MSF.TextMatrix(18, 5) = ""
    MSF.TextMatrix(18, 7) = ""
    MSF.TextMatrix(18, 8) = ""
    
    MSF.TextMatrix(19, 1) = "19"
    MSF.TextMatrix(19, 2) = "Rechazo a modo común ECG"
    MSF.TextMatrix(19, 3) = ""
    MSF.TextMatrix(19, 4) = ""
    MSF.TextMatrix(19, 5) = ""
    MSF.TextMatrix(19, 7) = ""
    MSF.TextMatrix(19, 8) = ""
    
    MSF.TextMatrix(20, 1) = "20"
    MSF.TextMatrix(20, 2) = "Filtro 50Hz ECG"
    MSF.TextMatrix(20, 3) = ""
    MSF.TextMatrix(20, 4) = ""
    MSF.TextMatrix(20, 5) = ""
    MSF.TextMatrix(20, 7) = ""
    MSF.TextMatrix(20, 8) = ""
    
    MSF.TextMatrix(21, 1) = "21"
    MSF.TextMatrix(21, 2) = "Offset y Ganancia ECG (electrodos)"
    MSF.TextMatrix(21, 3) = ""
    MSF.TextMatrix(21, 4) = ""
    MSF.TextMatrix(21, 5) = ""
    MSF.TextMatrix(21, 7) = ""
    MSF.TextMatrix(21, 8) = ""
    
    MSF.TextMatrix(22, 1) = "22"
    MSF.TextMatrix(22, 2) = "Constante de tiempo ECG (electrodos)"
    MSF.TextMatrix(22, 3) = ""
    MSF.TextMatrix(22, 4) = ""
    MSF.TextMatrix(22, 5) = ""
    MSF.TextMatrix(22, 7) = ""
    MSF.TextMatrix(22, 8) = ""
    
    MSF.TextMatrix(23, 1) = "23"
    MSF.TextMatrix(23, 2) = "Offset y Ganancia ECG (paletas)"
    MSF.TextMatrix(23, 3) = ""
    MSF.TextMatrix(23, 4) = ""
    MSF.TextMatrix(23, 5) = ""
    MSF.TextMatrix(23, 7) = ""
    MSF.TextMatrix(23, 8) = ""


    MSF.TextMatrix(24, 1) = "24"
    MSF.TextMatrix(24, 2) = "Constante de tiempo ECG (paletas)"
    MSF.TextMatrix(24, 3) = ""
    MSF.TextMatrix(24, 4) = ""
    MSF.TextMatrix(24, 5) = ""
    MSF.TextMatrix(24, 7) = ""
    MSF.TextMatrix(24, 8) = ""

    MSF.TextMatrix(25, 1) = "25"
    MSF.TextMatrix(25, 2) = "Alarmas"
    MSF.TextMatrix(25, 3) = ""
    MSF.TextMatrix(25, 4) = ""
    MSF.TextMatrix(25, 5) = ""
    MSF.TextMatrix(25, 7) = ""
    MSF.TextMatrix(25, 8) = ""

    MSF.TextMatrix(26, 1) = "26"
    MSF.TextMatrix(26, 2) = "Derivaciones ECG"
    MSF.TextMatrix(26, 3) = ""
    MSF.TextMatrix(26, 4) = ""
    MSF.TextMatrix(26, 5) = ""
    MSF.TextMatrix(26, 7) = ""
    MSF.TextMatrix(26, 8) = ""

    MSF.TextMatrix(27, 1) = "27"
    MSF.TextMatrix(27, 2) = "Control de amplitud"
    MSF.TextMatrix(27, 3) = ""
    MSF.TextMatrix(27, 4) = ""
    MSF.TextMatrix(27, 5) = ""
    MSF.TextMatrix(27, 7) = ""
    MSF.TextMatrix(27, 8) = ""

    MSF.TextMatrix(28, 1) = "28"
    MSF.TextMatrix(28, 2) = "Control de volumen"
    MSF.TextMatrix(28, 3) = ""
    MSF.TextMatrix(28, 4) = ""
    MSF.TextMatrix(28, 5) = ""
    MSF.TextMatrix(28, 7) = ""
    MSF.TextMatrix(28, 8) = ""

    MSF.TextMatrix(29, 1) = "29"
    MSF.TextMatrix(29, 2) = "Congelamiento de imagen"
    MSF.TextMatrix(29, 3) = ""
    MSF.TextMatrix(29, 4) = ""
    MSF.TextMatrix(29, 5) = ""
    MSF.TextMatrix(29, 7) = ""
    MSF.TextMatrix(29, 8) = ""
    
    MSF.TextMatrix(30, 1) = "30"
    MSF.TextMatrix(30, 2) = "Frecuencia Cardíaca"
    MSF.TextMatrix(30, 3) = ""
    MSF.TextMatrix(30, 4) = ""
    MSF.TextMatrix(30, 5) = ""
    MSF.TextMatrix(30, 7) = ""
    MSF.TextMatrix(30, 8) = ""

    MSF.TextMatrix(31, 1) = "31"
    MSF.TextMatrix(31, 2) = "Llaves y teclas"
    MSF.TextMatrix(31, 3) = ""
    MSF.TextMatrix(31, 4) = ""
    MSF.TextMatrix(31, 5) = ""
    MSF.TextMatrix(31, 7) = ""
    MSF.TextMatrix(31, 8) = ""


    MSF.TextMatrix(32, 1) = "32"
    MSF.TextMatrix(32, 2) = "Control de Contraste "
    MSF.TextMatrix(32, 3) = ""
    MSF.TextMatrix(32, 4) = ""
    MSF.TextMatrix(32, 5) = ""
    MSF.TextMatrix(32, 7) = ""
    MSF.TextMatrix(32, 8) = ""


    MSF.TextMatrix(33, 1) = "33"
    MSF.TextMatrix(33, 2) = "Autonomía de uso con batería interna"
    MSF.TextMatrix(33, 3) = ""
    MSF.TextMatrix(33, 4) = ""
    MSF.TextMatrix(33, 5) = ""
    MSF.TextMatrix(33, 7) = ""
    MSF.TextMatrix(33, 8) = ""

    MSF.TextMatrix(34, 1) = "34"
    MSF.TextMatrix(34, 2) = "Demora en modo sincronizado"
    MSF.TextMatrix(34, 3) = ""
    MSF.TextMatrix(34, 4) = "ms"
    MSF.TextMatrix(34, 5) = ""
    MSF.TextMatrix(34, 6) = "VM <"
    MSF.TextMatrix(34, 7) = "40"
    MSF.TextMatrix(34, 8) = ""

    MSF.TextMatrix(35, 1) = "34"
    MSF.TextMatrix(35, 2) = "Presencia de espiga en QRS"
    MSF.TextMatrix(35, 3) = ""
    MSF.TextMatrix(35, 4) = ""
    MSF.TextMatrix(35, 5) = ""
    MSF.TextMatrix(35, 7) = ""
    MSF.TextMatrix(35, 8) = ""


    MSF.TextMatrix(36, 1) = "35"
    MSF.TextMatrix(36, 2) = "ECG desde paletas de desfibrilación"
    MSF.TextMatrix(36, 3) = ""
    MSF.TextMatrix(36, 4) = ""
    MSF.TextMatrix(36, 5) = ""
    MSF.TextMatrix(36, 7) = ""
    MSF.TextMatrix(36, 8) = ""

    
    MSF.TextMatrix(37, 1) = "36"
    MSF.TextMatrix(37, 2) = "Test interno"
    MSF.TextMatrix(37, 3) = ""
    MSF.TextMatrix(37, 4) = ""
    MSF.TextMatrix(37, 5) = ""
    MSF.TextMatrix(37, 7) = ""
    MSF.TextMatrix(37, 8) = ""


    MSF.TextMatrix(38, 1) = "37"
    MSF.TextMatrix(38, 2) = "Reporte de desfibrilación (modelos 'R' solamente)"
    MSF.TextMatrix(38, 3) = ""
    MSF.TextMatrix(38, 4) = ""
    MSF.TextMatrix(38, 5) = ""
    MSF.TextMatrix(38, 7) = ""
    MSF.TextMatrix(38, 8) = ""

    
    MSF.TextMatrix(39, 1) = "38"
    MSF.TextMatrix(39, 2) = "Sumario de eventos (modelos 'R' solamente)"
    MSF.TextMatrix(39, 3) = ""
    MSF.TextMatrix(39, 4) = ""
    MSF.TextMatrix(39, 5) = ""
    MSF.TextMatrix(39, 7) = ""
    MSF.TextMatrix(39, 8) = ""

    
    MSF.TextMatrix(40, 1) = "39"
    MSF.TextMatrix(40, 2) = "Fecha y hora (modelos 'R' solamente)"
    MSF.TextMatrix(40, 3) = ""
    MSF.TextMatrix(40, 4) = ""
    MSF.TextMatrix(40, 5) = ""
    MSF.TextMatrix(40, 7) = ""
    MSF.TextMatrix(40, 8) = ""
    
    MSF.TextMatrix(41, 1) = "40"
    MSF.TextMatrix(41, 2) = "Oximetría (modelos 'S' solamente)"
    MSF.TextMatrix(41, 3) = ""
    MSF.TextMatrix(41, 4) = ""
    MSF.TextMatrix(41, 5) = ""
    MSF.TextMatrix(41, 7) = ""
    MSF.TextMatrix(41, 8) = ""
    
    MSF.TextMatrix(42, 1) = "41"
    MSF.TextMatrix(42, 2) = "Seguridad eléctrica (según IEC 60601-1)"
    MSF.TextMatrix(42, 3) = ""
    MSF.TextMatrix(42, 4) = ""
    MSF.TextMatrix(42, 5) = ""
    MSF.TextMatrix(42, 7) = ""
    MSF.TextMatrix(42, 8) = ""
    End Sub
Sub CARGA_REGISTRO_DATOS()

    Campos$ = "Item./F5;Descripción/A48;Rango/F6.2;Rango aceptable/F6.2;Rango aceptable/F6.2;Registro/F12.2;U./A6"
    Call CARGA_ENCABEZADO(MSF, Campos$, Me, , , , 7)
    
    MSF.MergeRow(0) = True
    MSF.MergeCells = flexMergeFree
    MSF.Rows = 47
    MSF.FixedRows = 1

    MSF.TextMatrix(1, 1) = "1"
    MSF.TextMatrix(1, 2) = "Ensambles y conexionados"
    
    MSF.TextMatrix(2, 1) = "2"
    MSF.TextMatrix(2, 2) = "Circuito de red de corriente alterna"
    
    MSF.TextMatrix(3, 1) = "3"
    MSF.TextMatrix(3, 2) = "Audio e indicadores luminosos"
    
    MSF.TextMatrix(4, 1) = "4"
    MSF.TextMatrix(4, 2) = "Configuración EEPROM"
    
    MSF.TextMatrix(5, 1) = "5"
    MSF.TextMatrix(5, 2) = "Fuente de alimentación general"
    
    MSF.TextMatrix(6, 1) = "6"
    MSF.TextMatrix(6, 2) = "Fuente de alimentación +6,8V"
    
    MSF.TextMatrix(7, 1) = "7"
    MSF.TextMatrix(7, 2) = "Fuente de alimentación +5V ECG"
    
    MSF.TextMatrix(8, 1) = "8"
    MSF.TextMatrix(8, 2) = "Fuente de alimentación +5V CPU"
    
    MSF.TextMatrix(9, 1) = "9"
    MSF.TextMatrix(9, 2) = "Fuente de alimentación +12V ECG"
    
    MSF.TextMatrix(10, 1) = "10"
    MSF.TextMatrix(10, 2) = "Entrega 2 Joules"
    MSF.TextMatrix(10, 3) = ""
    MSF.TextMatrix(10, 4) = "J"
    
    MSF.TextMatrix(11, 1) = "11"
    MSF.TextMatrix(11, 2) = "Entrega 10 Joules"
    MSF.TextMatrix(11, 3) = ""
    MSF.TextMatrix(11, 4) = "J"
    MSF.TextMatrix(11, 5) = "8.30"
    MSF.TextMatrix(11, 7) = "V"
    
    MSF.TextMatrix(12, 1) = "12"
    MSF.TextMatrix(12, 2) = "Entrega 50 Joules"
    MSF.TextMatrix(12, 3) = "-8.30"
    MSF.TextMatrix(12, 4) = "? VM ?"
    MSF.TextMatrix(12, 5) = "-7.70"
    MSF.TextMatrix(12, 7) = "V"
    
    MSF.TextMatrix(13, 1) = "13"
    MSF.TextMatrix(13, 2) = "Entrega 200 Joules"
    MSF.TextMatrix(13, 3) = "14.0"
    MSF.TextMatrix(13, 4) = "? VM ?"
    MSF.TextMatrix(13, 5) = "16.0"
    MSF.TextMatrix(13, 7) = "V"
    
    MSF.TextMatrix(14, 1) = "14"
    MSF.TextMatrix(14, 2) = "Entrega 300 Joules"
    MSF.TextMatrix(14, 3) = ""
    MSF.TextMatrix(14, 4) = "? VM ?"
    MSF.TextMatrix(14, 5) = "245"
    MSF.TextMatrix(14, 7) = "V"
    
    MSF.TextMatrix(15, 1) = "15"
    MSF.TextMatrix(15, 2) = "Entrega 360 Joules"
    MSF.TextMatrix(15, 3) = "7.0"
    MSF.TextMatrix(15, 4) = "? VM ?"
    MSF.TextMatrix(15, 5) = "8.0"
    MSF.TextMatrix(15, 7) = "V"
    
    MSF.TextMatrix(16, 1) = "16"
    MSF.TextMatrix(16, 2) = "Entrega interna 360 Joules (x10)"
    MSF.TextMatrix(16, 3) = "100"
    MSF.TextMatrix(16, 4) = "? VM ?"
    MSF.TextMatrix(16, 5) = "500"
    MSF.TextMatrix(16, 7) = "µA"
    
    MSF.TextMatrix(17, 1) = "17"
    MSF.TextMatrix(17, 2) = "Tiempo de carga a 360 Joules"
    MSF.TextMatrix(17, 3) = "0.2"
    MSF.TextMatrix(17, 4) = "? VM ?"
    MSF.TextMatrix(17, 5) = "0.5"
    MSF.TextMatrix(17, 7) = "A"
    
    MSF.TextMatrix(18, 1) = "18"
    MSF.TextMatrix(18, 2) = "Descarga automática de energía acumulada"
    MSF.TextMatrix(18, 3) = "0.1"
    MSF.TextMatrix(18, 4) = "? VM ?"
    MSF.TextMatrix(18, 5) = "3.5"
    MSF.TextMatrix(18, 7) = "A"
    
    MSF.TextMatrix(19, 1) = "19"
    MSF.TextMatrix(19, 2) = "Rechazo a modo común ECG"
    MSF.TextMatrix(19, 3) = "100"
    MSF.TextMatrix(19, 4) = "? VM ?"
    MSF.TextMatrix(19, 5) = "400"
    MSF.TextMatrix(19, 7) = "mVpp"
    
    MSF.TextMatrix(20, 1) = "20"
    MSF.TextMatrix(20, 2) = "Filtro 50Hz ECG"
    MSF.TextMatrix(20, 3) = "127"
    MSF.TextMatrix(20, 4) = "? VM ?"
    MSF.TextMatrix(20, 5) = "128"
    MSF.TextMatrix(20, 7) = "u"
    
    MSF.TextMatrix(21, 1) = "21"
    MSF.TextMatrix(21, 2) = "Offset y Ganancia ECG (electrodos)"
    MSF.TextMatrix(21, 3) = "221"
    MSF.TextMatrix(21, 4) = "? VM ?"
    MSF.TextMatrix(21, 5) = "223"
    MSF.TextMatrix(21, 7) = "u"
    
    MSF.TextMatrix(22, 1) = "22"
    MSF.TextMatrix(22, 2) = "Constante de tiempo ECG (electrodos)"
    MSF.TextMatrix(22, 3) = ""
    MSF.TextMatrix(22, 4) = ""
    MSF.TextMatrix(22, 5) = ""
    MSF.TextMatrix(22, 7) = ""
    
    MSF.TextMatrix(23, 1) = "23"
    MSF.TextMatrix(23, 2) = "Offset y Ganancia ECG (paletas)"
    MSF.TextMatrix(23, 3) = "285"
    MSF.TextMatrix(23, 4) = "? FC ?"
    MSF.TextMatrix(23, 5) = "295"
    MSF.TextMatrix(23, 7) = "ppm"

    MSF.TextMatrix(24, 1) = "24"
    MSF.TextMatrix(24, 2) = "Constante de tiempo ECG (paletas)"
    MSF.TextMatrix(24, 3) = ""
    MSF.TextMatrix(24, 4) = ""
    MSF.TextMatrix(24, 5) = ""
    MSF.TextMatrix(24, 7) = ""

    MSF.TextMatrix(25, 1) = "25"
    MSF.TextMatrix(25, 2) = "Alarmas"
    MSF.TextMatrix(25, 3) = ""
    MSF.TextMatrix(25, 4) = ""
    MSF.TextMatrix(25, 5) = ""
    MSF.TextMatrix(25, 7) = ""

    MSF.TextMatrix(26, 1) = "26"
    MSF.TextMatrix(26, 2) = "Derivaciones ECG"
    MSF.TextMatrix(26, 3) = ""
    MSF.TextMatrix(26, 4) = ""
    MSF.TextMatrix(26, 5) = ""
    MSF.TextMatrix(26, 7) = ""

    MSF.TextMatrix(27, 1) = "27"
    MSF.TextMatrix(27, 2) = "Control de amplitud"
    MSF.TextMatrix(27, 3) = ""
    MSF.TextMatrix(27, 4) = ""
    MSF.TextMatrix(27, 5) = ""
    MSF.TextMatrix(27, 7) = ""

    MSF.TextMatrix(28, 1) = "28"
    MSF.TextMatrix(28, 2) = "Control de volumen"
    MSF.TextMatrix(28, 3) = ""
    MSF.TextMatrix(28, 4) = ""
    MSF.TextMatrix(28, 5) = ""
    MSF.TextMatrix(28, 7) = ""

    MSF.TextMatrix(29, 1) = "29"
    MSF.TextMatrix(29, 2) = "Congelamiento de imagen"
    MSF.TextMatrix(29, 3) = ""
    MSF.TextMatrix(29, 4) = ""
    MSF.TextMatrix(29, 5) = ""
    MSF.TextMatrix(29, 7) = ""

    MSF.TextMatrix(30, 1) = "30"
    MSF.TextMatrix(30, 2) = "Frecuencia Cardíaca"
    MSF.TextMatrix(30, 3) = "4240"
    MSF.TextMatrix(30, 4) = "? VM ?"
    MSF.TextMatrix(30, 5) = "4260"
    MSF.TextMatrix(30, 7) = "V"

    MSF.TextMatrix(31, 1) = "31"
    MSF.TextMatrix(31, 2) = "Llaves y teclas"
    MSF.TextMatrix(31, 3) = ""
    MSF.TextMatrix(31, 4) = "VM = 253"
    MSF.TextMatrix(31, 5) = ""
    MSF.TextMatrix(31, 7) = "u"

    MSF.TextMatrix(32, 1) = "32"
    MSF.TextMatrix(32, 2) = "Control de Contraste "
    MSF.TextMatrix(32, 3) = ""
    MSF.TextMatrix(32, 4) = "VM ? 4370"
    MSF.TextMatrix(32, 5) = ""
    MSF.TextMatrix(32, 7) = "V"

    MSF.TextMatrix(33, 1) = "33"
    MSF.TextMatrix(33, 2) = "Autonomía de uso con batería interna"
    MSF.TextMatrix(33, 3) = ""
    MSF.TextMatrix(33, 4) = ""
    MSF.TextMatrix(33, 5) = ""
    MSF.TextMatrix(33, 7) = ""

    MSF.TextMatrix(34, 1) = "34"
    MSF.TextMatrix(34, 2) = "Demora en modo sincronizado"
    MSF.TextMatrix(34, 3) = ""
    MSF.TextMatrix(34, 4) = ""
    MSF.TextMatrix(34, 5) = ""
    MSF.TextMatrix(34, 7) = ""

    MSF.TextMatrix(35, 1) = "34"
    MSF.TextMatrix(35, 2) = "Presencia de espiga en QRS"
    MSF.TextMatrix(35, 3) = ""
    MSF.TextMatrix(35, 4) = ""
    MSF.TextMatrix(35, 5) = ""
    MSF.TextMatrix(35, 7) = ""

    MSF.TextMatrix(36, 1) = "35"
    MSF.TextMatrix(36, 2) = "ECG desde paletas de desfibrilación"
    MSF.TextMatrix(36, 3) = ""
    MSF.TextMatrix(36, 4) = ""
    MSF.TextMatrix(36, 5) = ""
    MSF.TextMatrix(36, 7) = ""
    
    MSF.TextMatrix(36, 1) = "36"
    MSF.TextMatrix(36, 2) = "Test interno"
    MSF.TextMatrix(36, 3) = ""
    MSF.TextMatrix(36, 4) = ""
    MSF.TextMatrix(36, 5) = ""
    MSF.TextMatrix(36, 7) = ""

    MSF.TextMatrix(37, 1) = "37"
    MSF.TextMatrix(37, 2) = "Reporte de desfibrilación (modelos 'R' solamente)"
    MSF.TextMatrix(37, 3) = "40"
    MSF.TextMatrix(37, 4) = "? VM ?"
    MSF.TextMatrix(37, 5) = "50"
    MSF.TextMatrix(37, 7) = "mV"
    
    MSF.TextMatrix(38, 1) = "38"
    MSF.TextMatrix(38, 2) = "Sumario de eventos (modelos 'R' solamente)"
    MSF.TextMatrix(38, 3) = "197.0"
    MSF.TextMatrix(38, 4) = "? VM ?"
    MSF.TextMatrix(38, 5) = "203.0"
    MSF.TextMatrix(38, 7) = "J"
    
    MSF.TextMatrix(39, 1) = "39"
    MSF.TextMatrix(39, 2) = "Fecha y hora (modelos 'R' solamente)"
    MSF.TextMatrix(39, 3) = "357.0"
    MSF.TextMatrix(39, 4) = "? VM ?"
    MSF.TextMatrix(39, 5) = "363.0"
    MSF.TextMatrix(39, 7) = "J"
    
    MSF.TextMatrix(40, 1) = "40"
    MSF.TextMatrix(40, 2) = "Oximetría (modelos 'S' solamente)"
    MSF.TextMatrix(40, 3) = ""
    MSF.TextMatrix(40, 4) = "VM ?"
    MSF.TextMatrix(40, 5) = "5.5"
    MSF.TextMatrix(40, 7) = "s"
    
    MSF.TextMatrix(41, 1) = "41"
    MSF.TextMatrix(41, 2) = "Seguridad eléctrica (según IEC 60601-1)"
    MSF.TextMatrix(41, 3) = ""
    MSF.TextMatrix(41, 4) = "VM ?"
    MSF.TextMatrix(41, 5) = "324"
    MSF.TextMatrix(41, 7) = "j"
    
End Sub








Function VALIDARRESULTADO$(MSF As msFlexGrid, REG&, COL&)
   COLUM% = COL&
   
   If Abs(XVALO(MSF.TextMatrix(REG&, COL&))) > 0 Then
        VALOR_a_VALIDAR = ROUND(XVALO(MSF.TextMatrix(REG&, COL&)), 2)
        VALOR1 = ROUND(XVALO(MSF.TextMatrix(REG&, 5)), 2)
        VALOR2 = ROUND(XVALO(MSF.TextMatrix(REG&, 7)), 2)
        SIMBOLEFT$ = Left(MSF.TextMatrix(REG&, 6), 1)
        SIMBORIGHT$ = Right(MSF.TextMatrix(REG&, 6), 1)
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
        End If
        
'
'
'        If VALOR_a_VALIDAR >= VALOR1 And VALOR_a_VALIDAR <= VALOR2 Then
'           VALIDARRESULTADO$ = "OK"
'           Call COLOREAR_TEXTO_SOLO_UNA_AFILA(MSF, vbBlack, REG&, COLUM%)
        Else
50         VALIDARRESULTADO$ = "MAL"
           Call COLOREAR_TEXTO_SOLO_UNA_AFILA(MSF, vbRed, REG&, COLUM%)
  End If
'        VALIDARRESULTADO$ = "OK"
'        Call COLOREAR_TEXTO_SOLO_UNA_AFILA(MSF, vbRed, REG&, COLUM%)
'  End If
  
End Function

Function VALIDARRESULTADOOKYMAL$(MSF As msFlexGrid, REG&, COL&)
   COLUM% = COL&
   
'   If Abs(XVALO(MSF.TextMatrix(REG&, COL&))) > 0 = "" Then
'      Exit Function
'   End If
   If TRIM$(MSF.TextMatrix(REG&, COL&)) = "" Then
      VALIDARRESULTADOOKYMAL$ = "OK"
   ElseIf MSF.TextMatrix(REG&, COL&) = "OK" Then
      VALIDARRESULTADOOKYMAL$ = "MAL"
   ElseIf MSF.TextMatrix(REG&, COL&) = "MAL" And REG& = 35 And Label15.Visible = True Then 'SI ES REPARACION Y FILA 35
      VALIDARRESULTADOOKYMAL$ = "N/A"
   ElseIf MSF.TextMatrix(REG&, COL&) = "N/A" And REG& = 35 And Label15.Visible = True Then 'SI ES REPARACION Y FILA 35
      VALIDARRESULTADOOKYMAL$ = ""
   Else
      VALIDARRESULTADOOKYMAL$ = ""
   End If
'
'
'        VALOR_a_VALIDAR = ROUND(XVALO(MSF.TextMatrix(REG&, COL&)), 2)
'        VALOR1 = ROUND(XVALO(MSF.TextMatrix(REG&, 5)), 2)
'        VALOR2 = ROUND(XVALO(MSF.TextMatrix(REG&, 7)), 2)
'        If VALOR_a_VALIDAR >= VALOR1 And VALOR_a_VALIDAR <= VALOR2 Then
'           VALIDARRESULTADO$ = "OK"
'           Call COLOREAR_TEXTO_SOLO_UNA_AFILA(MSF, vbBlack, REG&, COLUM%)
'
'        Else
'           VALIDARRESULTADO$ = "MAL"
'           Call COLOREAR_TEXTO_SOLO_UNA_AFILA(MSF, vbRed, REG&, COLUM%)
'        End If
'  End If

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


