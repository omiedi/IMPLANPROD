VERSION 5.00
Begin VB.Form GENUMSE07 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generador de Números de Serie"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   7020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00CDDADA&
      Height          =   615
      Left            =   120
      TabIndex        =   27
      Top             =   3480
      Visible         =   0   'False
      Width           =   5895
      Begin VB.CheckBox Check2 
         BackColor       =   &H00CDDADA&
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
         TabIndex        =   29
         Top             =   240
         Width           =   2715
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00CDDADA&
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
         Left            =   3360
         TabIndex        =   28
         Top             =   240
         Width           =   2475
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Excel 5.0;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   ""
      Top             =   1680
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00CDDADA&
      Caption         =   "Número de Serie Generado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   19
      Top             =   2520
      Width           =   5895
      Begin VB.Label Label8 
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
         Left            =   3360
         TabIndex        =   23
         Top             =   480
         Width           =   2460
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -1080
         TabIndex        =   22
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label10 
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
         Left            =   360
         TabIndex        =   21
         Top             =   480
         Width           =   2460
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2280
         TabIndex        =   20
         Top             =   240
         Width           =   1425
      End
   End
   Begin VB.CommandButton Command3 
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
      Left            =   5550
      TabIndex        =   8
      Top             =   480
      Width           =   175
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00400000&
      Height          =   5490
      Left            =   6120
      ScaleHeight     =   5430
      ScaleWidth      =   960
      TabIndex        =   3
      Top             =   -480
      Width           =   1020
      Begin VB.CommandButton Command4 
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
         Height          =   640
         Left            =   120
         Picture         =   "GENUMSE07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Carga de Números de Serie"
         Top             =   1800
         Width           =   650
      End
      Begin VB.CommandButton Command1 
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
         Height          =   650
         Left            =   120
         Picture         =   "GENUMSE07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Graba Numero de Serie"
         Top             =   2640
         Width           =   650
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
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
         Height          =   610
         Left            =   120
         Picture         =   "GENUMSE07.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   680
         Width           =   645
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   630
         TabIndex        =   4
         Top             =   3405
         Width           =   690
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   5
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "GENUMSE07.frx":075E
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   3600
         Width           =   1515
      End
   End
   Begin VB.TextBox Text12 
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
      Height          =   285
      Left            =   4440
      TabIndex        =   2
      Text            =   " "
      Top             =   480
      Width           =   1185
   End
   Begin VB.CommandButton Command9 
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
      Left            =   2835
      TabIndex        =   1
      ToolTipText     =   "Ordenes de Fabricación Todas"
      Top             =   480
      Visible         =   0   'False
      Width           =   175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CDDADA&
      Caption         =   "Generador de Números de Serie"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   5895
      Begin VB.CommandButton Command5 
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
         Left            =   2520
         TabIndex        =   30
         ToolTipText     =   "Solo Pendientes de Asignar N°. de Serie"
         Top             =   360
         Width           =   175
      End
      Begin VB.TextBox Text1 
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
         Height          =   285
         Left            =   4440
         MaxLength       =   6
         TabIndex        =   24
         Text            =   " "
         Top             =   750
         Width           =   1185
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "B.R.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3960
         TabIndex        =   32
         Top             =   1555
         Width           =   495
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
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
         Left            =   4560
         TabIndex        =   31
         Top             =   1510
         Width           =   1065
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tensión:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3600
         TabIndex        =   25
         Top             =   800
         Width           =   750
      End
      Begin VB.Label Label13 
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
         TabIndex        =   0
         Top             =   360
         Width           =   1545
      End
      Begin VB.Label Label24 
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
         Left            =   1200
         TabIndex        =   18
         Top             =   1920
         Width           =   1500
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dato Fijo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -780
         TabIndex        =   17
         Top             =   1965
         Width           =   1815
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
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
         Left            =   4560
         TabIndex        =   16
         Top             =   1920
         Width           =   1065
      End
      Begin VB.Label Label7 
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
         Height          =   225
         Left            =   2640
         TabIndex        =   15
         Top             =   1965
         Width           =   1815
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "F.Alta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3480
         TabIndex        =   14
         Top             =   400
         Width           =   750
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
         Left            =   -220
         TabIndex        =   13
         Top             =   400
         Width           =   1125
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Código / Modelo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -20
         TabIndex        =   12
         Top             =   800
         Width           =   1575
      End
      Begin VB.Label Label4 
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
         Left            =   1680
         TabIndex        =   11
         Top             =   750
         Width           =   1545
      End
      Begin VB.Label Label3 
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
         Left            =   240
         TabIndex        =   10
         Top             =   1140
         Width           =   5385
      End
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
   End
   Begin VB.Menu mnuConfiguracion 
      Caption         =   "Configuración"
      Begin VB.Menu MnuTablasAux 
         Caption         =   "Tablas Auxiliares"
      End
      Begin VB.Menu mnuFormatoNSerie 
         Caption         =   "Formato Números de Serie"
      End
   End
End
Attribute VB_Name = "GENUMSE07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Sub Command1_Click()
    'VALIDACIONES
    '
    '\\\OT-08-0929-OD-06/10/08///
    '
    If Label13 = "" Then
       Call MENSERR(24, "Falta Orden de Fabricación")
       Exit Sub
    End If
    
    VDEVV$ = TRIM$(Text12.TEXT)
    If FECHANUM(VDEVV$) < 1 Then
      Call MENSERR(24, "Fecha Inválida")
      Exit Sub
    End If
    '
    If InStr(UCase$(EMPREAC$), "MEDITEA") > 0 Then
        If TRIM(Text1) = "" Then
          Call MENSERR(24, "Falta Informar Tensión")
          Exit Sub
        End If
    End If
    '
    Screen.MousePointer = 11
    '
    Command1.Enabled = False
    Cantixxx& = XVALO(Label6)
    NORFA$ = Label13
    FE% = FECHANUM(Text12)
    FECHAALTA = FETEXCOR1$(FE%)
    CODINTER% = CODINT%(Label4)
    Descripcion$ = Label3
    MESYANXX$ = TRIM$(Label24)
    NumOrden% = MaxNumOr%(MESYANXX$)
    '
    'GRABACION EN TABLA TANUMSE
    
    SQLX$ = "SELECT * FROM TANUMSE WHERE COD_INTE=0"
    '
    UTILIZA_MODO_CONFIGURACION% = 0
    If CantRegistros("FORMANSERIE", "CORRESPONDE_A= 'NUMEROS DE SERIE' AND REFER_DATO = 'ORDEN' AND CANTICARACT > 0") > 0 Then
        X1X$ = PARTEFIJA$(CODINTER%, RET_TOTAL_CON_CODI$, FE%, OK%)
        If OK% >= 0 Then
          UTILIZA_MODO_CONFIGURACION% = 1
        Else
           UTILIZA_MODO_CONFIGURACION% = 0
        End If
    
    End If

    
    Dim TANUMTMP As ADODB.Recordset
    Set TANUMTMP = New ADODB.Recordset
    TANUMTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    For U& = 1 To Cantixxx&
        With TANUMTMP
            NumOrden% = NumOrden% + 1
            .AddNew   ' YA CONVERTIDO
            On Error Resume Next
            !CodiEmpr = CodiEmp%
            On Error GoTo 0
            !IDSUBOR = NORFA$
            !FechAlta = CVDAT(FE%)
            !Cod_Inte = CODINTER%
            !cod_exte = CODEXT$(CODINTER%)
            !Descrip = Descripcion$
            !MESANIO = TRIM$(MESYANXX$)
            !TEN_ENSA = XVALO(Text1)
              CAHASTA$ = TRIM(Str(NumOrden%))
            If UTILIZA_MODO_CONFIGURACION% > 0 Then
                 CAHASTA$ = FORMATOSERIEORD$(XVALO(CAHASTA$))
                 NUMSERIDES$ = FORMATOCONMASCARA$(RET_TOTAL_CON_CODI$ & CAHASTA$)
              Else
                 'MODO ORIGINAL MEDITEA
                 While Len(CAHASTA$) < 4: CAHASTA$ = "0" + CAHASTA$: Wend
                 NUMSERIDES$ = Right(CODEXT$(CODINTER%), 4) + MESYANXX$ + CAHASTA$
            End If
            !NUMEROSERIE = NUMSERIDES$
            !NUMEORD = NumOrden%
            !DepoNumSe = (-1)
            !NUMEBORE = XVALO(Label15)
            .Update
            '
        End With
        '

        'GRABACION EN EL ANOTADOR (HISTORIAL DE EQUIPOS)
        '
        CUERPO$ = "Alta N/Serie a Partir de: " & NORFA$
        Call ANOTANUSE(NUMSERIDES$, CUERPO$)

     Next U&
     TANUMTMP.Close
     Set TANUMTMP = Nothing
     
     If Control("CONFTRAZ", "ASIGLOTE") = "SI" Then
       Call AltaEnTrazaxSerie(NORFA$, CODINTER%, 1) 'agrega los mismos registros del tanumse a tabla trazaxserie
     End If
     
     
     Call ACTLISELNUMSER  'Actualiza lista de selecciona de numeros de serie
     Call ACTLISNUMSERXOF
     '
     MESYANXX$ = TRIM$(Label24)
     NumOrden% = MaxNumOr%(MESYANXX$) - Cantixxx&
     '
     NUMEORD_EYM% = XVALO(Right(Label10, 3))
     
     ALTOX% = ALTO_QR_NSERIE 'FUNCION
     ANCHOX% = ANCHO_QR_NSERIE 'FUNCION
     
     'LA CANTIDAD DE ETIQUETAS EN ESTE CASO NO SE TIENEN EN CUENTA POR QUE ES N.SERIE Y DEBERIA SER SIEMPRE 1
     CANT_ETIQ_X_FILAX = CANT_ETIQ_X_FILA: If CANT_ETIQ_X_FILAX < 1 Then CANT_ETIQ_X_FILAX = 1 'FUNCION
     '
     If InStr(UCase$(EMPREAC$), "MEDITEA") > 0 Then
        For U& = 1 To Cantixxx&
           NumOrden% = NumOrden% + 1
           CAHASTA$ = TRIM(Str(NumOrden%))
           While Len(CAHASTA$) < 4: CAHASTA$ = "0" + CAHASTA$: Wend
           NUMSERIDES$ = Right(CODEXT$(CODINTER%), 4) + MESYANXX$ + CAHASTA$
           Call MENUMSE07.PRINTETIQ(CODINTER%, NUMSERIDES$)
        Next U&
     Else
        Dim NSERIEX$(): Dim ORDEX$()
        For U& = 1 To Cantixxx&
           CAHASTA$ = TRIM(Str(NUMEORD_EYM%))
           CAHASTA$ = FORMATOSERIEORD$(XVALO(CAHASTA$))
           FILTX$ = NUMSERIDES$ & ";" & Right(FETEXCOR1$(FE%), 4)
           NUMSERIDES$ = RET_TOTAL_CON_CODI$ & CAHASTA$
           
           PIMPRE$ = Control$(IDENTER$, "PORTSTIK")
           If PIMPRE$ = "" Then PIMPRE$ = Control$("*", "PORTSTIK")
'           Call IMPRIMIR_ETIQUETAQR(CONTROL("", "STIKAPRO"), FILTX$, NUMSERIDES$, "IMAGNSERIE_QR.JPG", 1, "QR_NSERIE", PIMPRE$, ANCHOX%, ALTOX%)
           'ACA SOLO SE IMPRIMEN LAS ETIQUETAS PEQUEÑAS
           If Check3.Value = 1 Then
               ReDim Preserve NSERIEX$(U&) ' NUMERO DE SERIE
               NSERIEX$(U&) = NUMSERIDES$
           End If
           If Check2.Value = 1 Then
               ReDim Preserve ORDEX$(U&)
               ORDEX$(U&) = Right(NUMSERIDES$, 3) ' NUMERO DE ORDEN DE SERIES GENERADAS
           End If
           NUMEORD_EYM% = NUMEORD_EYM% + 1
        Next U&
     End If
     '
     FORIMPRE$ = Control$("", "STIKAPR2")
     k& = 0
     PARAMX1$ = "-"
     PARAMX2$ = "-"
     PARAMX3$ = "-"

     If Control("NUMSERIE", "IMPETIEX") = "SI" Then
       If Check2.Value > 0 Then
        CAITAB1% = 0
        If UBound(ORDEX$) > 0 Then
          For i& = 1 To UBound(ORDEX$)
'              NUMSERIDES$ = NSERIEX$(i&)
'              ORDEX$ = Right(NUMSERIDES$, 3)
              k& = k& + 1
              Select Case k&
                   Case 1: PARAMX1$ = AJUCEN$(ORDEX$(i&), 8)
                   Case 2: PARAMX2$ = AJUCEN$(ORDEX$(i&), 8)
                   Case 3: PARAMX3$ = AJUCEN$(ORDEX$(i&), 8)
              End Select
              If k& Mod 3 = 0 Then
                   GoSub IMPRIMIE_ETIQ
                   PARAMX1$ = "-"
                   PARAMX2$ = "-"
                   PARAMX3$ = "-"
                   k& = 0
              End If
            Next i&
            If k& > 0 Then GoSub IMPRIMIE_ETIQ
        End If
       End If
     End If
     '
     k& = 0
     PARAMX1$ = "-"
     PARAMX2$ = "-"
     PARAMX3$ = "-"
     If Check3.Value > 0 Then
        FORIMPRE$ = Control("", "STIKAPR3")
        If UBound(NSERIEX$) > 0 Then
          For i& = 1 To UBound(NSERIEX$)
              NUMSERIDES$ = FORMATOCONMASCARA$(NSERIEX$(i&))
              k& = k& + 1
              Select Case k&
                   Case 1: PARAMX1$ = AJUCEN$(NUMSERIDES$, 8)
                   Case 2: PARAMX2$ = AJUCEN$(NUMSERIDES$, 8)
                   Case 3: PARAMX3$ = AJUCEN$(NUMSERIDES$, 8)
              End Select
              If k& Mod 3 = 0 Then
                   GoSub IMPRIMIE_ETIQ
                   PARAMX1$ = "-"
                   PARAMX2$ = "-"
                   PARAMX3$ = "-"
                   k& = 0
              End If
           Next i&
           If k& > 0 Then GoSub IMPRIMIE_ETIQ
        End If
     End If
     
'     FORIMPRE$ = CONTROL$("", "STIKAPR2")
'     If CONTROL("NUMSERIE", "IMPETIEX") = "SI" Then
'       If Check2.Value > 0 Then
'        CAITAB1% = 0
'        If UBound(ORDEX$) > 0 Then
'          For i& = 1 To UBound(ORDEX$) Step 3
'             If i& = 1 Then
'                  Call LIMPIATETAB
'                  CODTABU1$(1) = "COD-CONJ"
'                  CODTABU1$(2) = "COD-MPRI"
'                  CODTABU1$(3) = "CPOS-CLI"
'                  CODTABU1$(4) = "DES-CONJ"
'                  CODTABU1$(5) = "REF-CONJ"
'                  CODTABU1$(6) = "UNI-CONJ"
'                  CACOLTAB1% = 6
'                  CAITAB1% = CAITAB1% + 1
'                  TETABU1$(4, CAITAB1%) = AJUCEN$(ORDEX$(i&), 8)
'                  If UBound(ORDEX$) >= i& + 1 Then TETABU1$(5, CAITAB1%) = AJUCEN$(ORDEX$(i& + 1), 8)
'                  If UBound(ORDEX$) >= i& + 2 Then TETABU1$(6, CAITAB1%) = AJUCEN$(ORDEX$(i& + 2), 8)
'             Else
'                  CAITAB1% = CAITAB1% + 1
'                  TETABU1$(1, CAITAB1%) = AJUCEN$(ORDEX$(i&), 8)
'                  If UBound(ORDEX$) >= i& + 1 Then TETABU1$(2, CAITAB1%) = AJUCEN$(ORDEX$(i& + 1), 8)
'                  If UBound(ORDEX$) >= i& + 2 Then TETABU1$(3, CAITAB1%) = AJUCEN$(ORDEX$(i& + 2), 8)
'             End If
'          Next i&
'           'IMPRIME
'           PUERIMPRE$ = CONTROL("*", "PRIETIQ1")
'           If CONTROL(IDENTER$, "PRIETIQ1") <> "" Then PUERIMPRE$ = CONTROL(IDENTER$, "PRIETIQ1")
'           Call SETPRINPORT(PUERIMPRE$)
'           Call PRINTDOC("@" + FORIMPRE$)
'           Call SETPRINPORT("RESTORE")
'        End If
'       End If
'       If Check3.Value > 0 Then
'        CAITAB1% = 0
'        If UBound(NSERIEX$) > 0 Then
'          For i& = 1 To UBound(NSERIEX$) Step 3
'             If i& = 1 Then
'                  Call LIMPIATETAB
'                  CODTABU1$(1) = "COD-CONJ"
'                  CODTABU1$(2) = "COD-MPRI"
'                  CODTABU1$(3) = "CPOS-CLI"
'                  CODTABU1$(4) = "DES-CONJ"
'                  CODTABU1$(5) = "REF-CONJ"
'                  CODTABU1$(6) = "UNI-CONJ"
'                  CACOLTAB1% = 6
'                  CAITAB1% = CAITAB1% + 1
'                  TETABU1$(4, CAITAB1%) = FORMATOCONMASCARA$(NSERIEX$(i&))
'                  If UBound(ORDEX$) >= i& + 1 Then TETABU1$(5, CAITAB1%) = FORMATOCONMASCARA$(NSERIEX$(i& + 1))
'                  If UBound(ORDEX$) >= i& + 2 Then TETABU1$(6, CAITAB1%) = FORMATOCONMASCARA$(NSERIEX$(i& + 2))
'             Else
'                  CAITAB1% = CAITAB1% + 1
'                  TETABU1$(1, CAITAB1%) = FORMATOCONMASCARA$(NSERIEX$(i&))
'                  If UBound(NSERIEX$) >= i& + 1 Then TETABU1$(2, CAITAB1%) = FORMATOCONMASCARA$(NSERIEX$(i& + 1))
'                  If UBound(NSERIEX$) >= i& + 2 Then TETABU1$(3, CAITAB1%) = FORMATOCONMASCARA$(NSERIEX$(i& + 2))
'             End If
'           Next i&
'           'IMPRIME
'           PUERIMPRE$ = CONTROL("*", "PRIETIQ1")
'           If CONTROL(IDENTER$, "PRIETIQ1") <> "" Then PUERIMPRE$ = CONTROL(IDENTER$, "PRIETIQ1")
'           Call SETPRINPORT(PUERIMPRE$)
'
'           Call PRINTDOC("@" + FORIMPRE$)
'           Call SETPRINPORT("RESTORE")
'        End If
'       End If
'
'     End If
     '
     If Control("NUMSERIE", "UTIARMEX") <> "NO" Then
         'ASIGNACION DE ARMADOR
         Dim obj As New Art_Importados
         CODI$ = Label4.Caption
         Call obj.VeriImportado(CODI$) 'VERIFICA QUE NO SEA IMPORTADO
         If obj.EsImportado < 1 Then
           If COMALTER%("Desea Asignar Ahora Un Armador Para Esta Orden\\No,Continuar\Sí, Asignar") = 2 Then
            Call GRABAARMADOR(NORFA$)
           End If
         End If
         Set obj = Nothing
     End If
     
     '
     Call BLANFORMU
     '
     Screen.MousePointer = 1
     '
99   Command1.Enabled = True
Exit Sub


IMPRIMIE_ETIQ:
        
        PUERIMPRE$ = Control("*", "PRIETIQ1")
        If Control(IDENTER$, "PRIETIQ1") <> "" Then PUERIMPRE$ = Control(IDENTER$, "PRIETIQ1")
        Call SETPRINPORT(PUERIMPRE$)
        FILTX$ = PARAMX1$ & ";" & PARAMX2$ & ";" & PARAMX3$
        
        Call STDFORM(FORIMPRE$, FILTX$, "PRINT")

        Return
        
      
90      Call SETPRINPORT("RESTORE")
Return
     '\\\OT-08-0929-OD-06/10/08///
End Sub

Function ALTO_QR_NSERIE()
   Condi$ = " CORRESPONDE_A= 'NUMEROS DE SERIE' AND REFER_DATO = 'ALTO'"
   ALTO_QR_NSERIE = XVALO(VALOBADA("FORMANSERIE", "MEDIDAS", Condi$, "NOMESS"))

End Function

Function CANT_ETIQ_X_FILA()
   CANT_ETIQ_X_FILA = XVALO(Control("ORDEFABR", "CANTCOLU")) 'CONFIGURACION EN SETUP DE OF.
End Function
Private Sub Command2_Click()
  '\\\OT-08-0929-OD-06/10/08///
   Hide
  '\\\OT-08-0929-OD-FIN///
End Sub

Private Sub Command3_Click()
   '\\\OT-08-0929-OD-06/10/08///
    Call CommSel("SELFECHA", Command3, Text12)
    '\\\OT-08-0929-OD-FIN///
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    '
    If XVALO(Control("CONFTRAZ", "ASIGLOTE")) > 0 Then GoTo 99
    
5   RTA1% = ALTERNA%("Carga Directa de Números de Serie\Importación de Números de Serie\Carga Directa de N/Serie (Sin Mov. de Stock)\Baja de Nº Serie")
    If RTA1% < 1 Or RTA1% > 4 Then GoTo 99
    '
    If RTA1% = 1 Then
       If DERACCE%("CADINUSE", "Carga Directa de Numeros de Serie") >= 2 Then
          GoTo 10
       End If
    End If
    '
    If RTA1% = 3 Then
       If DERACCE%("CADINUSE", "Carga Directa de N/Serie (Sin Mov. de Stock)") >= 2 Then
          GoTo 10
       End If
    End If
    '
    If RTA1% = 2 Then
       If DERACCE%("IMPONUSE", "Importacion de Numeros de Serie") >= 2 Then
          GoTo 120
       End If
    End If
    '
    If RTA1% = 4 Then
       If DERACCE%("BAJANUSE", "Baja de Nº Serie") >= 2 Then
          GoTo 60 ' baja de numeros de serie
       End If
    End If
    GoTo 99
    '
10  TTXX$ = "Esta Opcion ha sido Transitoriamente Deshabilitada.\   \Solo puede Realizarse el Ingreso Manual\de Equipos que Quedaron en Depósito -1\con Orden de Fabricación Cerrada"
    Call COMUNI(TTXX$)     ' 6 DE MAYO 2010 - epb
    GoTo 30
    '
    '
    Call REPLA$(VDEF$, MKI$(HOY(HO$)), 35, 2) ' PRESENTA POR DEFECTO LA FECHA ACTUAL
    '
    OBX$ = OBJPLA$("CARNUMSE", VDEF$) ' OBJETO TANUMSE PARA INGRESAR LOS NUMEROS DE SERIE
    VDEF$ = OBX$
    If OBX$ = "" Then GoTo 5
    NUSERIEE$ = TRIM$(Mid$(VDEF$, 1, 32))
    CI% = CVI(Mid$(VDEF$, 33, 2))
    fech% = CVI(Mid$(VDEF$, 35, 2))
    FECC = FECHATEX(fech%)
    TENSION$ = Mid$(VDEF$, 37, 6)
    NuDepo% = CVI(Mid$(VDEF$, 43, 2))

    CANTIRE& = CantRegistros&(" TANUMSE", "NumeroSerie = '" & NUSERIEE$ & "'") 'VERIFICA SI EL NUMERO DE SERIE EXISTE
    '
    If CANTIRE& > 0 Then Call MENSERR(24, "Número de Serie Existente"): GoTo 10 ' VALIDACIONES
    If CI% < 1 Then Call MENSERR(24, "Código no Válido"): GoTo 10
    '
    If InStr(EMPREAC$, "MEDITEA") > 0 Then
      If Len(NUSERIEE$) > 4 Then
        Ulti4Cod$ = Right(CODEXT$(CI%), 4) ' Valida que los ultimos cuatro numeros del codigo coincidan con los 4 primeros del numero de serie
        Prim4NuSe1$ = Left$(NUSERIEE$, 4)
       '
       If Ulti4Cod$ <> Prim4NuSe1$ Then
         Call COMUNI("Número de Serie Difiere de Código")
         GoTo 10
       End If
      Else
        Call MENSERR(24, "Número de Serie Cantidad de Caracteres menor a 5"): GoTo 10
      End If
    End If
    '
    Deposito$ = ECOARCH$("TADEPOSI", Chr$(NuDepo%))
    If Deposito$ = "" Then Call MENSERR(24, "Depósito no Válido"): GoTo 10
    '
    MESYANIO$ = Mid$(NUSERIEE$, 5, 4)
    NORDEN% = XVALO(Mid$(NUSERIEE$, 9, 4))
    '
    FECHAA$ = FECHATEX$(fech%)
    CMOV$ = TRIM$(Control$("", "CODAJUST")) ' SE OBTIENE PROXIMO NUMERO DE AJUSTE
    If CMOV$ = "" Then CMOV$ = "AJ"
    NVM$ = TRIM$(ProNroComprobante("AJ"))
    While Len(NVM$) < 6: NVM$ = "0" + NVM$: Wend
    NRO& = XVALO(NVM$)
    If CMOV$ = "" Then CMOV$ = "AJ"
    DOSEC$ = "AJ." + Mid$(FECHAA$, 7, 2) + Mid$(FECHAA$, 4, 2) + Left$(FECHAA$, 2)
    DOPRI$ = CMOV$ + "." + NVM$
    REFE$ = "AJUSTE POR CARGA DIRECTA"
    DEPX% = NuDepo%

    SQLX$ = "SELECT * FROM TANUMSE WHERE COD_INTE=0" ' SI PASO LAS VALIDACIONES AGREGA EL NUEVO NUMERO DE SERIE
    Dim TANUTMP As ADODB.Recordset
    Set TANUTMP = New ADODB.Recordset
    TANUTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
      With TANUTMP
        .AddNew   ' YA CONVERTIDO
        On Error Resume Next
        !CodiEmpr = CodiEmp%
        On Error GoTo 0
        !NUMEROSERIE = NUSERIEE$
        !MESANIO = MESYANIO$
        !NUMEORD = NORDEN%
        !Cod_Inte = CI%
        !cod_exte = CODEXT(CI%)
        Descripcion$ = DESCRIT(CI%)
        !Descrip = Descripcion$
        !FechAlta = FECC
        !TEN_ENSA = XVALO(TENSION$)
        !DepoNumSe = NuDepo%
        .Update
        '
        If RTA1% = 1 Then Call MOVISTOK(fech%, CI%, NUSERIEE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NORDEN%, 0, 1, NC%, DEPX%, 1)
        '
      End With
    TANUTMP.Close
    Set TANUTMP = Nothing
    '
    Call ACTLISELNUMSER   'Actualiza lista de selecciona de numeros de serie
    '
    'GRABACION EN EL ANOTADOR (HISTORIAL DE EQUIPOS)
    '
    CUERPO$ = "Alta N/Serie en Forma Manual"
    Call ANOTANUSE(NUSERIEE$, CUERPO$)
    '

    Call PRINTALBANUMSE(NUSERIEE$, CI%, FECC, 1)
    '
    Call MENSERR(24, "Alta Realizada")
    VDEF$ = ""
    Screen.MousePointer = 1
    GoTo 10
    '
    '
30  ' ingreso de equipos en deposito -1
    If DERACCE%("INGREQUI", "Ingreso de Equipos del Deposito -1") < 2 Then
       GoTo 5
    End If
    '
    Call CARGALISEL("NUSEMEN1", "TANUMSE", "NUMEROSERIE/A16;COD_INTE/C16 + MASTER/A32;IDSUBOR/A14", "NUMEROSERIE <> '' AND DEPONUMSE < 0 ORDER BY NUMEROSERIE ASC")
    Call SELECHO("NUSEMEN1")
    NSX$ = TRIM$(VALACT1$("NUSEMEN1"))
    If NSX$ <> "" Then
       CIXI% = CODINT%(Left$(VALACT2$("NUSEMEN1"), 16))
       NOFA$ = TRIM$(VALOBADA("TANUMSE", "IDSUBOR", "NUMEROSERIE = " & NSX$ & " AND COD_INTE = " & CIXI%, "NOMESS"))
       If NOFA$ = "" Then
          Call MENSERR(24, "Numero de O/Fabricacion '" + NOFA$ + "' Inexistente o no Valido")
          GoTo 30
       End If
       '
       Condi$ = "IDSUBOR = '" & NOFA$ & "' "
       STORFA = XVALO(VALOBADA("ORDEFABR", "STATUORF", Condi$, "NOMESS"))
       CANAP = XVALO(VALOBADA("ORDEFABR", "CANAPRO", Condi$, "NOMESS"))
       CANPR = XVALO(VALOBADA("ORDEFABR", "CANPROY", Condi$, "NOMESS"))
       If CANAP >= CANPR Then STORFA = 3
       If STORFA <> 3 Then
          Call MENSERR(24, "Numero de O/Fabricacion '" + NOFA$ + "' no Está Cerrada")
          GoTo 30
       End If
       '
       Call SELECHO("TADEPOSI")
       DEPOEN = XVALO(VALACT1$("TADEPOSI"))
       If DEPOEN < 1 Or DEPOEN >= 99 Then GoTo 30
       '
       TTXX$ = "¿ Confirma el Ingreso del Equipo"
       TTXX$ = TTXX$ + "\" + CODEXT$(CIXI%) + " - " + DESCRIT0$(CIXI%)
       TTXX$ = TTXX$ + "\Numero de Serie: " + NSX$
       TTXX$ = TTXX$ + "\al Deposito Numero " & DEPOEN & " ?"
       If COMALTER%(TTXX$ + "\\Si, Ingresar\No, Cancelar") = 1 Then
          Call ACTUREGISTRO("TANUMSE", "DEPONUMSE", "NUMEROSERIE = '" & NSX$ & "' AND COD_INTE = " & CIXI%, DEPOEN, RAFE&)
          CUERPO$ = "Ingreso Directo N/Serie del -1 al " & DEPOEN
          Call ANOTANUSE(NSX$, CUERPO$)
       End If
       GoTo 30
    End If
    GoTo 5
    '
60  Call SELECHO("MASTER")
    CODXX$ = TRIM$(VALACT1$("MASTER"))
    If CODXX$ = "" Then GoTo 5
    CIXI% = CODINT%(CODXX$)
    '
'    Call COPYSTRU("INDINUSE", "INDINUMS") ' BAJA DE NUMEROS DE SERIE 'OD-02/12/11 SE UTILIZA SOLO LOCAL
    TTXX$ = REGBLAN$("INDINUSE")
'    Call SETULTREG("!INDINUMS", 0)''OD-02/12/11 NO HACE FALTA POR QUE LUEGO QUE GRABA HACE SETULTREG
    Screen.MousePointer = 11
    '
    CONDI1$ = "(NumRemito = '' OR NumRemito IS NULL) AND Cod_inte=" & CIXI% 'OD-02/12/11 TODO A CONDI
    FIN& = CantRegistros&("TANUMSE", CONDI1$)
    '
    SQLX$ = "SELECT NUMEROSERIE,COD_INTE,COD_EXTE FROM TANUMSE WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE " & CONDI1$
    SQLX$ = SQLX$ + " ORDER BY COD_EXTE, NUMEROSERIE ASC "
    Dim NUSERTMP As ADODB.Recordset
    Set NUSERTMP = READSET(SQLX$)
    '
    With NUSERTMP
     JJ& = 0
     Do While Not .EOF
      Call TRACE(20, JJ&, FIN&)
      NUSERR$ = SAFETEXT$(!NUMEROSERIE)
      NusePed$ = VALOBADA("PEDDETA", "NUMEROSERIE", "NUMEROSERIE = '" & NUSERR$ & "'")
      If NusePed$ = "" Then
          CI1% = XVALO(!Cod_Inte)
          Call REPLA(TTXX$, NUSERR$, 1, 16)
          Call REPLA(TTXX$, CODEXT$(CI1%) + "  " + DESCRIT(CI1%), 17, 47)
          JJ& = JJ& + 1
          Call GRAREG("!INDINUSE", TTXX$, JJ&)
      End If
      .MoveNext
     Loop
    End With
    NUSERTMP.Close
    Set NUSERTMP = Nothing
    Call SETULTREG("!INDINUSE", JJ&)
'    Call TRACENTRAL("INDINUMS", "") 'OD-02/12/11
    Screen.MousePointer = 1
    If ULTREG&("!INDINUSE") < 1 Then
     Call MENSERR(24, "No Hay Números de Serie que se Puedan Eliminar")
     GoTo 5
    End If
    '
70  Call SELECHO("!INDINUSE")
    NSR$ = TRIM$(VALACT1$("!INDINUSE"))
    If NSR$ = "" Then GoTo 5
       '
    If COMALTER("Realmente Desea dar de Baja el Número de Serie: " & NSR$ & " \\Cancelar\Dar de Baja") <> 2 Then GoTo 70
    Screen.MousePointer = 11
    SQLX$ = " IdenLote = '" & NSR$ & "'"
    SQLX$ = SQLX$ + " AND CodiMovi = 'AJ'"
    LOT$ = VALOBADA("MOVISTO", "IdenLote", SQLX$)
    CI% = XVALO(VALOBADA("TANUMSE", "Cod_Inte", "NUMEROSERIE = '" & NSR$ & "'")) ' CARGAMOS DATOS PARA FUNCION MOVISTO E IMPRESION DE ETIQUETA DE BAJA
    FF% = HOY(HO$)
    FECHAX$ = FECHATEX$(FF%)
    If LOT$ <> "" Then ' SI EXISTE EN LA TABLA MOVISTO REALIZAMOS MOVIMIENTO INVERESO
      DEP% = XVALO(VALOBADA("TANUMSE", "DepoNumse", "NUMEROSERIE = '" & NSR$ & "'"))
      CMOV$ = TRIM$(Control$("", "CODAJUST")) ' SE OBTIENE PROXIMO NUMERO DE AJUSTE
      If CMOV$ = "" Then CMOV$ = "AJ"
      NVM$ = TRIM$(ProNroComprobante("AJ"))
      While Len(NVM$) < 6: NVM$ = "0" + NVM$: Wend
      DOSEC$ = "AJ." + Mid$(FECHAX$, 7, 2) + Mid$(FECHAX$, 4, 2) + Left$(FECHAX$, 2)
      DOPRI$ = CMOV$ + "." + NVM$
      REFE$ = "Anul.S/N.: " & NSR$
      Call MOVISTOK(FF%, CI%, NSR$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NORDEN%, 0, 1, NC%, DEP%, (-1))
    End If
    '
    SQLJ$ = "DELETE FROM TANUMSE"
    SQLJ$ = SQLJ$ + " WHERE NUMEROSERIE = '" & NSR$ & "'"
    FLEXCONN.Execute (FILTSQL$(SQLJ$))
    Call MENSERR(24, "Número de Serie: " & NSR$ & " Dado de Baja")
    Call ACTLISELNUMSER   'Actualiza lista de selecciona de numeros de serie
    '
    Call PRINTALBANUMSE(NSR$, CI%, FECHAX$, 2)
    '
    Screen.MousePointer = 1
    GoTo 99 ' FIN BAJA DE NUMEROS DE SERIE
    '
120 CommonDialog1.CancelError = True ' SELECCION DE ARCHIVO A IMPORTAR Y GRABACION
    CommonDialog1.DialogTitle = "Archivo Origen de Datos"
    '     cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
    CommonDialog1.FILTER = "Hojas de Cálculo Excel(*.XLS)|*.XLS|"
    CommonDialog1.FilterIndex = 1
    On Error GoTo 5
    CommonDialog1.ShowOpen
    On Error GoTo 0

    DoEvents
    Screen.MousePointer = 11
    Data1.DatabaseName = UCase$(CommonDialog1.FileName)
    On Error Resume Next
    Data1.Refresh
    Data1.RecordSource = Data1.Database.TableDefs(0).Name
    Data1.Refresh
    On Error GoTo 0
    '
    SQLX$ = "SELECT * FROM TANUMSE"
    Dim TANUMTMP As ADODB.Recordset
    Set TANUMTMP = New ADODB.Recordset
    TANUMTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    With Data1
     .Recordset.MoveFirst
     Do While Not .Recordset.EOF
        NUSERIE$ = ""
        On Error Resume Next
        NUSERIE$ = .Recordset.Fields(0)
        CANTIRE& = CantRegistros&(" TANUMSE", "NumeroSerie = '" & NUSERIE$ & "'") 'VERIFICA SI EL NUMERO DE SERIE EXISTE
        If CANTIRE& > 0 Then GoTo 20
        On Error GoTo 0
        '
        Codigo$ = ""
        On Error Resume Next
        Codigo$ = .Recordset.Fields(1)
        CI% = CODINT(Codigo$)
        If CI% < 1 Then
          Call MENSERR(24, "Código '" & Codigo$ & "' Inexistente Dar de Alta\y Volver a Correr el Proceso de Importación")
          GoTo 20
        End If
        On Error GoTo 0
        '
        FEALTA = ""
        On Error Resume Next
        FEALTA = .Recordset.Fields(2)
        On Error GoTo 0
        '
        TENSION$ = ""
        On Error Resume Next
        TENSION$ = .Recordset.Fields(3)
        On Error GoTo 0
        '
        NUDEP$ = 0
        On Error Resume Next
        NUDEP$ = .Recordset.Fields(4)
        On Error GoTo 0
        NuDepo% = XVALO(NUDEP$)
        Deposito$ = ECOARCH$("TADEPOSI", Chr$(NuDepo%))
        If Deposito$ = "" Then
          Call MENSERR(24, "Número de Depósito:  '" & NuDepo% & "' Inválido dar de Alta \y Volver a Correr el Proceso de Importación")
          GoTo 20
        End If
        MESYANIO$ = Mid$(NUSERIE$, 5, 4)
        NORDEN% = XVALO(Mid$(NUSERIE$, 9, 4))
        '
        FECH1% = CVINT(FEALTA)
        FECHAA$ = FECHATEX$(FECH1%)
        DIA$ = Mid$(FECHAA$, 1, 2)
        MES$ = Mid$(FECHAA$, 4, 2)
        AÑO$ = Mid$(FECHAA$, 7, 2)
        FECHAAA$ = AÑO$ + "-" + MES$ + "-" + DIA$
        '
        CMOV$ = "AJ"
        DOPRI$ = "AJ-" + FECHAAA$
        REFE$ = "AJUSTE INICIAL POR IMPORTACION"
        DEPX% = NuDepo%
        '
        TANUMTMP.AddNew ' AGREGO A LA TABLE TANUMSE LOS DATOS LEIDOS   ' YA CONVERTIDO
                On Error Resume Next
        TANUMTMP!CodiEmpr = CodiEmp%
                On Error GoTo 0
        TANUMTMP!MESANIO = MESYANIO$
        TANUMTMP!NUMEORD = NORDEN%
        TANUMTMP!Cod_Inte = CI%
        Descripcion$ = DESCRIT0$(CI%)
        TANUMTMP!Descrip = Descripcion$
        TANUMTMP!TEN_ENSA = XVALO(TENSION$)
        TANUMTMP!FechAlta = CVDAT(CVINT(FEALTA))
        TANUMTMP!NUMEROSERIE = NUSERIE$
        TANUMTMP!cod_exte = Codigo$
        TANUMTMP!DepoNumSe = NuDepo%
        TANUMTMP.Update
        '
        Call MOVISTOK(fech%, CI%, NUSERIE$, CMOV$, DOPRI$, DOPRI$, DOPRI$, DOPRI$, REFE$, NORDEN%, 0, 1, NC%, DEPX%, 1)
        '
20      .Recordset.MoveNext
        
      Loop
     End With
     TANUMTMP.Close
     Set TANUMTMP = Nothing
     '
     Call ACTLISELNUMSER   'Actualiza lista de selecciona de numeros de serie
     '
     Screen.MousePointer = 1
     Call MENSERR(24, "Proceso de Importación Finalizada")
     '
99  Command4.Enabled = True
    
End Sub

Private Sub Command5_Click()
   campos$ = "O.Fabric./A12;REFERENCIA/A64;Cant./F8.1;N°.BR./F8;Proveedor/A32"
   Call CARGA_ENCABEZADO(FRMZELECHO.msf2, campos$, FRMZELECHO)
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, campos$, FRMZELECHO)
   FRMZELECHO.Caption = "Ordenes de Fabricación Pendientes de Generar N°. de Serie"
 
   Dim obj As New RECORXET
   Set rs = obj.RS_OF_PENDIENTES_NSERIE_CON_BR()
   Call Carga_MSF_Recordset(rs, campos$, FRMZELECHO.msf2, 1)
   Set obj = Nothing
   rs.Close
   Set rs = Nothing

   FRMZELECHO.Width = 13000
   FRMZELECHO.Show 1
   If TRIM$(RESP_ZELE_VECT(1)) <> "" Then
     Label13 = TRIM$(RESP_ZELE_VECT(1))
     Label15 = TRIM$(RESP_ZELE_VECT(4))
   End If

End Sub

'Private Sub Command5_Click()
'   Command5.Enabled = False
'   '
'   SQLX$ = "SELECT * FROM TANUMSE WHERE LEN(MESANIO) < 4"
'   '
'   Dim TANUMTMP As ADODB.Recordset
'   Set TANUMTMP = New ADODB.Recordset
'   TANUMTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
'
'   With TANUMTMP
'     If Not (.EOF And .BOF) Then
'       .MoveFirst
'       Do While Not .EOF
'         NSERIE$ = SAFETEXT$(!NUMEROSERIE)
'         '
'         MESA$ = TRIM$(SAFETEXT$(!MESANIO))
'         While Len(MESA$) < 4: MESA$ = "0" + MESA$: Wend
'
'         NORDEN$ = TRIM$(FORMATNUM(XVALO(!NUMEORD), "I4"))
'         While Len(NORDEN$) < 4: NORDEN$ = "0" & NORDEN$: Wend
'         '
'         SERIE$ = Mid$(NSERIE$, 1, 4)
'         NUEVALOR$ = SERIE$ & MESA$ & NORDEN$
'         '
'         !MESANIO = MESA$
'         !NUMEROSERIE = NUEVALOR$
'         .Update
'         Call ACTUREGISTRO("PEDDETA", "NUMEROSERIE", "NUMEROSERIE = '" & NSERIE$ & "'", NUEVALOR$, REG&)
'         .MoveNext
'       Loop
'       TANUMTMP.Close
'       Set TANUMTMP = Nothing
'     End If
'   End With
'   '
'   Call COMUNI("Proceso de Recuperacion Finalizado Correctamente")
'   '
'   Command5.Enabled = True
'End Sub

Private Sub Command9_Click()
    '\\\OT-08-0929-OD-06/10/08///
    Call CommSel("ECORDEF/Indice de Ordenes de Fabricación", Command9, Label13)
    '\\\OT-08-0929-OD-FIN///
End Sub

Private Sub Form_Load()
    '\\\OT-08-0929-OD-06/10/08///
    Call ABREORFAB
    FF% = HOY(HO$)
    Text12.TEXT = HO$
    '\\\OT-08-0929-OD-FIN///
    '
    If Control("ORDEFABR", "IMETIQOF") = "SI" Then
       Label14.Visible = False
       Text1.Visible = False
    End If
    Call CREATABLA_STOCKS
    
    If Control("NUMSERIE", "IMPETIEX") = "SI" Then
       Frame3.Visible = True
       Check2.Value = 1
       Check3.Value = 1
       Me.Height = 4830
    End If
    
    '
End Sub

Private Sub Label13_Change()
    '\\\OT-08-0929-OD-06/10/08///
     '
     'ATENCION !!! PARA QUE ESTO FUNCIONE EN MEDITEA DEBE ESTAR CONFIGURADO EL FORMATO DE NUMERO DE SERIE DESDE ESTA
     'MISMA PANTALLA ENTRANDO CON EMPRESA MEDITEA VER MUESTRA DE FORMATO
     'LABEL DE ORDEN DE FABRICACION
     NORFA$ = TRIM$(Label13)
     If Label13 <> "" Then
        CANTIRE& = CantRegistros&(" TANUMSE", "IdSubOr = '" & NORFA$ & "'") 'verifica que no se haya ingresado antes
        If CANTIRE& < 1 Then 'si no encontro nada es poe que no es existente
            CIXI% = XVALO(VALOBADA("ORDEFABR", "Cod_Inte", "IdSubOr = '" & NORFA$ & "'"))
            '
            If CIXI% < 1 Then
               Call MENSERR(24, "Orden de Fabricación no Válida")
            End If
            '
            If ESTRAZABLE(CIXI%) <> 2 Then ' Verifica que el codigo sea trazable
              Call MENSERR(24, "Código No Trazable") ' si no es trazable 2 limpia y se va
              Call BLANFORMU
              Exit Sub
            End If
            '
            Label4 = CODEXT(CIXI%)
            
            'A PARTIR DE AQUI ACTUALIZACION POR NUEVO MODO DE CONFIGURACION DE NUMEROS DE SERIE
            If CantRegistros&("FORMANSERIE", "CORRESPONDE_A= 'NUMEROS DE SERIE'", "NOMESS") > 0 Then
               PARTEFIJAX$ = PARTEFIJA$(CIXI%, TOTAL_CON_CODI$, FECHANUM(Text12), OK%)
               If OK% >= 0 Then
               'ACA VALIDAR SI FORMATEA PARA EYM Y LUEGO TERMINAR AQUI CON EXIT SUB
               End If
            End If
            
'FALTA GENERAR UNA FUNCION QUE FORMATEE CON LA MASCARA
            '
            
            Label24 = PARTEFIJAX$
           
            'CODIGO ANTERIOR
            KODI$ = Right(Label4, 4)
            '
            If Len(KODI$) < 4 Then ' valida que el codigo no sea menor a 4 caracteres
             Call MENSERR(24, "Código no Admitido")
             Call BLANFORMU
             Exit Sub
            End If
            '
            Label3 = DESCRIT0$(CIXI%)
            Cantidad# = XVALO(VALOBADA("ORDEFABR", "CanProy", "IdSubOr = '" & NORFA$ & "'"))
            Label6 = FORMATNUM$(Cantidad#, "F8.0")
            Call DESDEHASTA(TOTAL_CON_CODI$)

            'Call ACTUALMESAÑO
            Text1 = "220"
            'APLICA LA TENSION POR DEFECTO SEGUN EL CODIGO
            'EN EL CASO QUE EL CODIGO EMPIECE CON 11 ES A 110 Y CON 10 ES A 220
             If Left$(TRIM$(Label4), 2) = "11" Then
                 Text1 = "110"
             Else
                 Text1 = "220"
             End If
            'fin  de cambio

        Else
           Call MENSERR(24, "Numeros de Serie Ya Asignados para la Orden de Fabricacion")
           Call BLANFORMU
        End If
     End If
     
     '\\\OT-08-0929-OD-FIN///

End Sub

'Private Sub Label13_Change() VERSION ANTERIOR VER EL TEMA DEL LABEL 24
'    '\\\OT-08-0929-OD-06/10/08///
'     '
'     'LABEL DE ORDEN DE FABRICACION
'     NORFA$ = TRIM$(Label13)
'     If Label13 <> "" Then
'        CANTIRE& = CantRegistros&(" TANUMSE", "IdSubOr = '" & NORFA$ & "'") 'verifica que no se haya ingresado antes
'        If CANTIRE& < 1 Then 'si no encontro nada es poe que no es existente
'            CIXI% = XVALO(VALOBADA("ORDEFABR", "Cod_Inte", "IdSubOr = '" & NORFA$ & "'"))
'            '
'            If CIXI% < 1 Then
'               Call MENSERR(24, "Orden de Fabricación no Válida")
'            End If
'            '
'            If ESTRAZABLE(CIXI%) <> 2 Then ' Verifica que el codigo sea trazable
'              Call MENSERR(24, "Código No Trazable") ' si no es trazable 2 limpia y se va
'              Call BLANFORMU
'              Exit Sub
'            End If
'            '
'            Label4 = CODEXT(CIXI%)
'            KODI$ = Right(Label4, 4)
'            '
'            If Len(KODI$) < 4 Then ' valida que el codigo no sea menor a 4 caracteres
'             Call MENSERR(24, "Código no Admitido")
'             Call BLANFORMU
'             Exit Sub
'            End If
'            '
'            Label3 = DESCRIT0$(CIXI%)
'            Cantidad# = XVALO(VALOBADA("ORDEFABR", "CanProy", "IdSubOr = '" & NORFA$ & "'"))
'            Label6 = FORMATNUM$(Cantidad#, "F8.0")
'            Call DESDEHASTA
'            Call ACTUALMESAÑO
'            Text1 = "220"
'            'APLICA LA TENSION POR DEFECTO SEGUN EL CODIGO
'            'EN EL CASO QUE EL CODIGO EMPIECE CON 11 ES A 110 Y CON 10 ES A 220
'             If Left$(TRIM$(Label4), 2) = "11" Then
'                 Text1 = "110"
'             Else
'                 Text1 = "220"
'             End If
'            'fin  de cambio
'
'        Else
'           Call MENSERR(24, "Numeros de Serie Ya Asignados para la Orden de Fabricacion")
'           Call BLANFORMU
'        End If
'     End If
'
'     '\\\OT-08-0929-OD-FIN///
'
'End Sub
Private Sub Label24_Change()
    '\\\OT-08-0929-OD-06/10/08///
    'LABEL DEL MESAÑO
'     If Label13 <> "" Then
'       Call ACTUALMESAÑO
'     End If
    '\\\OT-08-0929-OD-FIN///

End Sub


Private Sub Label6_Change()
    '\\\OT-08-0929-OD-06/10/08///
    'LABEL DE CANTIDAD
     If Label13 <> "" Then
       Call ACTUALMESAÑO
     End If
     '\\\OT-08-0929-OD-FIN///

End Sub




Private Sub mnuFormatoNSerie_Click()
    If CLACONTRA% < 1 Then Exit Sub
    
    VENTANA.Inicializar = 0
    VENTANA.CampEditables = "1;3;4;5;6;7;8"
    VENTANA.ColEcoCampo(1) = "CAMPO1;CAMPO2;DATASQL;LISTA = 'TIPODATO'"
    VENTANA.AgregaRegistro = 1
    VENTANA.UtilizaFormula = 1
    VENTANA.HabilitaBorrar = 1
    VENTANA.HabilitaInsertar = 1
    VENTANA.NoMostrarBotonTilde = 0
    VENTANA.TITULO = ""
    VENTANA.Inicializar = 1
    campos$ = "COD./F6;T.DATO/A24;Q.CARACT/F8;A PARTIR/F8;EXTREMO/A12;COMP./A12;MASCARA/A24;MEDIDA/F9,2"
    VENTANA.Campox = campos$
    
    Dim obj As New RECORXET
    Set RS1 = obj.RS_CONFIGURACION_DE_FORMATO_DE_NSERIE
    Call Carga_MSF_Recordset(RS1, campos$, VENTABNEW.MSF)
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    VENTABNEW.Show 1
    If VENTABNEW.APROBADO% > 0 Then
        Dim MSF As MSFlexGrid
        Set MSF = VENTABNEW.MSF
        
        
        Condi$ = "CORRESPONDE_A = 'NUMEROS DE SERIE'"
        Call ACTUREGISTRO("FORMANSERIE", "MARBORRA", Condi$, 1, REGAF&, "NOMESS")
        
        SQLX$ = " SELECT * FROM FORMANSERIE WHERE CORRESPONDE_A = 'NUMEROS DE SERIE'"
        Dim rs As ADODB.Recordset
        Set rs = New ADODB.Recordset
25      On Error Resume Next
        rs.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        If Err Then
         On Error GoTo 0
         Call CapturaErrorOpen
         GoTo 25
        End If
          With rs
              ORD% = 0
              For REG& = 1 To MSF.Rows - 1
                  .AddNew
                  ORD% = ORD% + 1
                  !TIP_DATO = MSF.TextMatrix(REG&, 1)
                  !REFER_DATO = MSF.TextMatrix(REG&, 2)
                  !CANTICARACT = MSF.TextMatrix(REG&, 3)
                  !APARTIR = MSF.TextMatrix(REG&, 4)
                  !EXTREMO = MSF.TextMatrix(REG&, 5)
                  !COMPLETACON = MSF.TextMatrix(REG&, 6)
                  !CORRESPONDE_A = "NUMEROS DE SERIE"
                  !MASCARA = MSF.TextMatrix(REG&, 7)
                  !MEDIDAS = XVALO(MSF.TextMatrix(REG&, 8))
                  !MARBORRA = 0
                  !ORDEN = ORD%
                  .Update
              Next REG&
          End With
        
          rs.Close
          Set rs = Nothing
          Condi$ = Condi$ + " AND  MARBORRA = 1  "
          Call BORRAREGISTROS("FORMANSERIE", Condi$, REGAF&, "NOMESS")
      End If
      Set VENTANA = Nothing

End Sub

Private Sub MnuTablasAux_Click()
  Call GENERAR_ESTRUCTURAS_SQL
  Call CARGA_ENCABEZADO(FRMZELECHO.msf2, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO)
  Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO)
  FRMZELECHO.Caption = "LISTA DE ESTRUCTURAS PRF Y RFS"

  Dim obj As New RECORXET
  Set RS1 = obj.RS_ENCAPRFLISTAS()
  Call Carga_MSF_Recordset(RS1, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO.msf2)
  Set obj = Nothing
  RS1.Close
  Set RS1 = Nothing
  EPAC$ = UCase$(TRIM$(EMPREAC$))
  '

  FRMZELECHO.Width = 8500
  FRMZELECHO.Show 1
  IDLIST& = XVALO(RESP_ZELE_VECT(1))
  If InStr(EPAC$, "ENVAPLAST") > 0 Then
    LIXTA$ = TRIM$(RESP_ZELE_VECT(2))
    If LIXTA$ = "MATERIALES ENVAPLAST" Then
       Call COMUNI("Atención !!! Editar Tabla de Familias Desde Maestro de Artículos")
       Exit Sub
    End If
  End If

  Call ABM_TABLA(IDLIST&)


End Sub

Private Sub Text12_Change()
   '\\\OT-08-0929-OD-06/10/08///
    If Label13 <> "" Then 'LABEL DE LA ORDEN DE FABRICACION
        Call Label13_Change
'        Call ACTUALMESAÑO
'        Call DESDEHASTA(Label24.Caption)
    End If
    '\\\OT-08-0929-OD-FIN///

End Sub

Private Sub Text12_DblClick()
   '\\\OT-08-0929-OD-06/10/08///
   Command3_Click
   '\\\OT-08-0929-OD-FIN///

End Sub
 Sub BLANFORMU()
    'LIMPIA TODOS LOS CONTROLES MENOS LA FECHA
    '\\\OT-08-0929-OD-06/10/08///
    Text1 = ""
    Label13 = ""
    Label4 = ""
    Label3 = ""
    Label24 = ""
    Label6 = ""
    Label10 = ""
    Label8 = ""
    Label15 = ""
    '\\\OT-08-0929-OD-FIN//
 End Sub
 
 
 Sub DESDEHASTA(TOTAL_CON_CODI$)
    '\\\OT-08-0929-OD-06/10/08///
    'ACTUALIZA LOS LABEL DESDE Y HASTA DE LOS NUMEROS DE SERIE GENERADOS
    VDEVV$ = TRIM$(Text12.TEXT)
    If FECHANUM(VDEVV$) > 0 Then ' SI LA FECHA ES VALIDA
      If Label13 <> "" Then      'SI HAY UNA ORDEN DE FABRICACION SELECCIONADA
         CANTIX$ = TRIM$(Str(1 + MaxNumOr%(Label24))) 'TOMA EL VALOR MAXIMO DE NUMERO DE ORDEN + 1
'         While Len(CANTIX$) < 4: CANTIX$ = "0" & CANTIX$: Wend ' FORMATEA
         CANTIX$ = FORMATOSERIEORD$(XVALO(CANTIX$))
         Label10 = FORMATOCONMASCARA$(TOTAL_CON_CODI$ & CANTIX$)
'         KODI$ = Right(Label4, 4)
'         Label10 = KODI$ & Label24 & CANTIX$ ' CONCATENA
         CANTHAST% = XVALO(CANTIX$) + (XVALO(Label6) - 1) 'CALCULA EL HASTA
         CAHASTA$ = TRIM$(Str(CANTHAST%))           'SUMANDO EL DESDE + LA CANTIDAD
'         While Len(CAHASTA$) < 4: CAHASTA$ = "0" + CAHASTA$: Wend
         Label8 = FORMATOCONMASCARA$(TOTAL_CON_CODI$ & FORMATOSERIEORD$(XVALO(CAHASTA$)))
      End If
     Else
      Label10 = ""
      Label8 = ""
    End If
    '\\\OT-08-0929-OD-FIN///

 End Sub
 '
 Sub ACTUALMESAÑO()
    '\\\OT-08-0929-OD-06/10/08///
'    'FORMATEA EL MES Y AÑO EN 4 DIGITOS Y LO ACTUALIZA
'    MESYANO$ = REPLACAR$(Text12, "/", "")
'    Label24 = Mid$(TRIM$(MESYANO$), 3, 4)
    '\\\OT-08-0929-OD-FIN///
 End Sub
 
 Sub ACTLISELNUMSER() ' Actualiza listado de seleccion de numeros de serie
    '
    Call CARGALISEL("INDINUSE", "TANUMSE", "NUMEROSERIE/A16;COD_EXTE/A18;DESCRIP/A24", "NUMEROSERIE <> '' ORDER BY FECHALTA ASC, NUMEORD ASC")
    '
 End Sub
 Sub ACTLISNUMSERXOF()
     Call CARGALISEL("INNUSEOF", "TANUMSE", "IDSUBOR/A16;COD_EXTE/A18;DESCRIP/A24", "NUMEROSERIE <> '' GROUP BY IDSUBOR ,COD_EXTE, DESCRIP ORDER BY IDSUBOR ASC")
 End Sub
 
 Sub PRINTALBANUMSE(NSERIE$, CI%, FECHA, Optional TIP%)
 
    FORIPRE$ = TRIM$(Control$("", "ALBANSE"))
    If FORIPRE$ = "" Then Exit Sub
    '
    Codigo$ = CODEXT$(CI%)
    DESXRIP$ = DESCRIT0$(CI%)
    NUSERIE$ = NSERIE$
    FECHEMI$ = FECHA
    '
    If TIP% = 1 Then
        NUMEDOC$ = NUSERIE$
        TIPODOC$ = "Alta de Numero de Serie"
    ElseIf TIP% = 2 Then
        NUMEDOC$ = NUSERIE$
        TIPODOC$ = "Baja de Numero de Serie"
    End If
    CODPARAM$(1) = "COD-MAT": DOCPARAM$(1) = Codigo$
    CODPARAM$(2) = "DES-MAT": DOCPARAM$(2) = DESXRIP$
    CODPARAM$(3) = "REF-MAT1": DOCPARAM$(3) = NUSERIE$
    CODPARAM$(4) = "FECH-EMI": DOCPARAM$(4) = FECHEMI$
    CODPARAM$(5) = "TIPO-COM": DOCPARAM$(5) = TIPODOC$
    CAPARDOC = 5
    '
    Call PRINTDOC("@" + FORIPRE$)
    '

 End Sub
