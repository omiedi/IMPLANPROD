VERSION 5.00
Begin VB.Form APROCLIB04 
   BackColor       =   &H00E0F0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aprobación de Material de Texto Libre"
   ClientHeight    =   6390
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   8505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   8505
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture15 
      Height          =   6400
      Left            =   7560
      ScaleHeight     =   6345
      ScaleWidth      =   1035
      TabIndex        =   36
      Top             =   0
      Visible         =   0   'False
      Width           =   1100
      Begin VB.Image Image3 
         Height          =   6400
         Left            =   0
         Picture         =   "APROCLIB04.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0F0E0&
      Caption         =   "Result"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   7560
      TabIndex        =   33
      Top             =   4640
      Width           =   1335
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
         Left            =   160
         Picture         =   "APROCLIB04.frx":2398
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Aprueba la Recepcion Provisoria."
         Top             =   945
         Width           =   650
      End
      Begin VB.CommandButton BOTEXIT 
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
         Left            =   160
         Picture         =   "APROCLIB04.frx":26A2
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Rechazo del Material."
         Top             =   315
         Width           =   650
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0F0E0&
      Height          =   735
      Left            =   4905
      TabIndex        =   29
      Top             =   4650
      Width           =   2535
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
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
         Height          =   285
         Left            =   860
         TabIndex        =   31
         Top             =   240
         Width           =   1485
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Remito:"
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
         Left            =   -1200
         TabIndex        =   30
         Top             =   310
         Width           =   1920
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0F0E0&
      Height          =   735
      Left            =   225
      TabIndex        =   24
      Top             =   4650
      Width           =   4575
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo:"
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
         Left            =   1440
         TabIndex        =   28
         Top             =   310
         Width           =   1920
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
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
         Height          =   285
         Left            =   3450
         TabIndex        =   27
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "E.V.:"
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
         Left            =   -1290
         TabIndex        =   26
         Top             =   310
         Width           =   1920
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
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
         Height          =   285
         Left            =   720
         TabIndex        =   25
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.PictureBox Picture11 
      BackColor       =   &H00800000&
      ForeColor       =   &H00000000&
      Height          =   4680
      Left            =   7560
      ScaleHeight     =   4620
      ScaleWidth      =   1005
      TabIndex        =   16
      Top             =   -105
      Width           =   1065
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
         Height          =   645
         Left            =   140
         Picture         =   "APROCLIB04.frx":29AC
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Cancela - Cierra la Aplicación"
         Top             =   420
         Width           =   645
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   140
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   17
         Top             =   3885
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -120
            TabIndex        =   18
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -210
         Picture         =   "APROCLIB04.frx":2AF6
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   1410
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0F0E0&
      Height          =   855
      Left            =   3345
      TabIndex        =   13
      Top             =   5475
      Width           =   4095
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
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
         Height          =   285
         Left            =   1745
         TabIndex        =   23
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Deposito Destino:"
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
         Left            =   -240
         TabIndex        =   22
         Top             =   420
         Width           =   1920
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0F0E0&
      Height          =   855
      Left            =   225
      TabIndex        =   12
      Top             =   5475
      Width           =   3015
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad Recibida:"
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
         Left            =   -110
         TabIndex        =   15
         Top             =   420
         Width           =   1920
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
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
         Height          =   285
         Left            =   1920
         TabIndex        =   14
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0F0E0&
      Height          =   3255
      Left            =   225
      TabIndex        =   4
      Top             =   1305
      Width           =   7215
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
         Height          =   2415
         Left            =   180
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   630
         Width           =   6855
      End
      Begin VB.TextBox Text1 
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
         Left            =   1140
         TabIndex        =   5
         Top             =   240
         Width           =   5895
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Material:"
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
         Left            =   -50
         TabIndex        =   11
         Top             =   360
         Width           =   960
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0F0E0&
      Height          =   1005
      Left            =   225
      TabIndex        =   0
      Top             =   210
      Width           =   7215
      Begin VB.Label Label14 
         BackColor       =   &H00C0C0FF&
         Height          =   375
         Left            =   4320
         TabIndex        =   21
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFC0&
         Height          =   375
         Left            =   3480
         TabIndex        =   20
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
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
         Height          =   255
         Left            =   2520
         TabIndex        =   19
         Top             =   360
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
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
         Height          =   285
         Left            =   1200
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nº O/C:"
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
         Left            =   120
         TabIndex        =   9
         Top             =   330
         Width           =   960
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
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
         Height          =   285
         Left            =   5955
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
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
         Left            =   4920
         TabIndex        =   7
         Top             =   330
         Width           =   960
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
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
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Top             =   615
         Width           =   975
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Proveedor:"
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
         Left            =   120
         TabIndex        =   2
         Top             =   690
         Width           =   960
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
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
         Height          =   285
         Left            =   2400
         TabIndex        =   1
         Top             =   615
         Width           =   4650
      End
   End
   Begin VB.Menu menu 
      Caption         =   "Menú"
      Begin VB.Menu cancelarechazo 
         Caption         =   "Cancelar Rechazo"
      End
      Begin VB.Menu aprobarrechazo 
         Caption         =   "Aprobar Rechazo"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "APROCLIB04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub aprobarrechazo_Click()
   Call Command2_Click
End Sub

Private Sub BOTEXIT_Click()
    BOTEXIT.Enabled = False
    '
    If COMALTER%("¿ Está Seguro de Rechazar la Partida ?\\No, Cancelar Rechazo\Si, Confirmar Rechazo") <> 2 Then
      GoTo 99
    End If
    '
    'ACA RECHAZA LA RECEPCION PROVISORIA.
    'EDICION  EN BOLRECEP
    Call ABRESTOCKS
    With Bolrecep
      '
      NBoRE& = Val(Label9)
      CARECI# = Val(Label7)
      SQLX$ = " NumeBoRe = " & NBoRE& & " "
      .FindFirst SQLX$
      If .NoMatch = False Then
        .Edit
          !CanToApro = 0
          !CanToRech = CARECI#
          !Status = 1 'APROBADA
          !UsuApro = USERNAME
          !FechLibe = Date
          !UsuApro = USERNAME   'NUEVO
          !FeHoApro = Now      'NUEVO
          !Observa = USERNAME 'nuevo
       .Update
      End If
      '
    End With
    '
    ' aqui re-activa la Orden de Compra por el rechazado mas el faltante
    RECHFAL = CARECI#
    If RECHFAL > 0.005 Then
      NOCORI& = Val(Label4)
      If NOCORI& > 0 Then
        Call ABRECOMPRAS
        ABUS$ = "Nume_Ocom = " & NOCORI& & " "
        ABUS$ = ABUS$ + "AND (Cant_Rec+0.005) >= " & Abs(RECHFAL)
        With OcomDeta
          .FindFirst ABUS$
          If .NoMatch = False Then
            .Edit
            !Cant_rec = !Cant_rec - RECHFAL
            !Stat_Ocom = 0
            .Update
          End If
        End With
      End If
    End If
    '
    '3)ACA FALTA HACER LA IMPRESION DEL REMITO POR RECHAZO
    Call REMRECHLIB
    '
    Call COMUNI("Se ha Procesado el Rechazo de la B.R. " + TRIM$(Str$(NBoRE)))
    Unload Me
    '
99 BOTEXIT.Enabled = True
End Sub

Sub REMRECHLIB() ' (NPROVE%, DEPX%, REMIRECH$)
   '
   FORIPRE$ = TRIM$(Control$("", "REMIRECH"))
   If FORIPRE$ = "" Then Exit Sub
   '
   If FORIPRE$ <> "" Then
      LU$ = LUDAT$
      If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
         RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
         Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Remito por Rechazo:\  \'" + RMCC$ + "'.")
         Exit Sub
      End If
   End If
   '
   Screen.MousePointer = 11
   '
   REMIRECH$ = "RD." + TRIM$(Label9)
   NC% = Abs(Val(Label5))
   DESTIDOC% = (-1) * NC%
   '
   FEX = HOY(HOS$)
   FECHDOC$ = FECHATEX$(FEX)
   NUMEDOC$ = REMIRECH$
   TIPODOC$ = "Remito de Devolución"
   '
   DTP$ = OBJPLA$("TRANSPOR", DATRANSPO$)
   DATRANSPO$ = DTP$
   '
   CODPARAM$(1) = "TRANSPOR": DOCPARAM$(1) = TRIM$(Mid$(DTP$, 1, 32))
   CODPARAM$(2) = "DOMI-TRA": DOCPARAM$(2) = TRIM$(Mid$(DTP$, 33, 32))
   CODPARAM$(3) = "CUIT-TRA": DOCPARAM$(3) = TRIM$(Mid$(DTP$, 65, 16))
   CODPARAM$(4) = "PATE-TRA": DOCPARAM$(4) = TRIM$(Mid$(DTP$, 81, 16))
   CODPARAM$(5) = "COND-TRA": DOCPARAM$(5) = TRIM$(Mid$(DTP$, 97, 32))
   CODPARAM$(6) = "DOCUCOND": DOCPARAM$(6) = TRIM$(Mid$(DTP$, 129, 24))
   CAPARDOC% = 6
   '
   CODTABU1$(1) = "COD-MAT"
   CODTABU1$(2) = "DES-MAT"
   CODTABU1$(3) = "CANTIDAD"
   CODTABU1$(4) = "CAKILOS"
   CACOLTAB1% = 4
   '
   AINDI% = 0
   CAITAB1% = 0
   '
   CAITAB1% = CAITAB1% + 1
   TETABU1$(1, CAITAB1%) = ""
   TETABU1$(2, CAITAB1%) = Text1
   TETABU1$(3, CAITAB1%) = Label7
   '
   FORDESCRI$ = TRIM$(UCase$(ATRIFORM$(FORIPRE$, "DES-MAT/FORMATO")))
   ANCHODES% = Val(Mid$(FORDESCRI$, 2))
   If ANCHODES% < 16 Then ANCHODES% = 16
   '
   CAITAB1% = CAITAB1% + 1
   TETABU1$(2, CAITAB1%) = "   " ' PARA DEJAR UN RENGLON EN BLANCO
   '
   TXTX$ = Text2
   Call FRATEXTO(TXTX$, ANCHODES%)
   '
   URN& = CARETEX%
   '
   For U& = 1 To CARETEX%
     TTXY$ = RTrim$(RETEX$(U&))
     CAITAB1% = CAITAB1% + 1
     TETABU1$(2, CAITAB1%) = TTYY$
   Next U&
   '
   Call PRINTDOC("REMIRECH")   ' Remito de Devolucion
   '
   Call CIERRARCH("*.*")
   Screen.MousePointer = 1
   '
End Sub

Private Sub cancelarechazo_Click()
    Call BOTEXIT_Click
End Sub

'
Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command2_Click()
    '
    Static CRR1
    'BUSCO Y MARCO NºB.R. EN LISTAMAT
    '
    NEV1$ = ""
    GRU1$ = ""
    CLI1$ = ""
    DEV1$ = ""
    DGR1$ = ""
    DCR1$ = ""
    PLA1$ = ""
    MAT1$ = ""
    LLA1$ = ""
    
    '
    NUOC& = Val(Label4)
    NUMEDO$ = NUOC&
    While Len(NUMEDO$) < 6: NUMEDO$ = "0" + NUMEDO$: Wend
    NUMEDO$ = "OC-" + NUMEDO$ + "-*"
    '
    NBoRE& = Val(Label9)
    Call ABREORTRA
    With ListaMat
        SQLX$ = " Nro_Ocom LIKE '" & NUMEDO$ & "' "
        .FindFirst SQLX$
        If .NoMatch = False Then
          '
          'FORMATEO EL NUMERO DE BR
          N1$ = CStr(NBoRE&)
          Do While Len(N1$) < 6
            N1$ = "0" + N1$
          Loop
          NBR2$ = "BR-" + N1$ + "-00"
          '
          NEV1$ = !Nro_Otrab
          GRU1$ = !Nro_Grup
          DCR1$ = !Descri_Cor
          PLA1$ = !NPlano
          MAT1$ = !Material
          LLA1$ = !CLAVE
          '
          .Edit
          !nro_bolre = NBR2$
          .Update
          '
        End If
    End With
    '
    If NEV1$ <> "" Then
      With OrdeTrab
        SQLX$ = " Nro_Otrab = '" & NEV1$ & "' "
        .FindFirst SQLX$
        If .NoMatch = False Then
          DEV1$ = !Refe_Otrab
          NC% = !Nclie_Otrab
          CLI1$ = RASOCLI$(NC%)
        End If
      End With
      '
      If GRU1$ <> "" Then
        With GrupOtra
          SQLX$ = " Nro_Otrab = '" & NEV1$ & "' "
          SQLX$ = SQLX$ + " AND Nro_Grup = '" & GRU1$ & "' "
          .FindFirst SQLX$
          If .NoMatch = False Then
            DGR1$ = !Deno_Grup
          End If
        End With
      End If
    End If
    '
    'EDICION  EN BOLRECEP
    Call ABRESTOCKS
    With Bolrecep
        '
        NUREPRO$ = TRIM$(Label22)
        CARECI# = Val(Label7)
        SQLX$ = " NumeBoRe = " & NBoRE& & " AND NumRem_Prov = '" & NUREPRO$ & "' "
        .FindFirst SQLX$
        If .NoMatch = False Then
          .Edit
            !CanToApro = CARECI#
            !CanToRech = 0
            !Status = 1 'APROBADA
            !UsuApro = USERNAME
            !Status = 1
            !FechLibe = Date
            !UsuApro = USERNAME   'NUEVO
            !FeHoApro = Now      'NUEVO
            !Observa = USERNAME 'nuevo
         .Update
        End If
        '
    End With
    '
    CAN1$ = Label7
    CRR1 = CRR1 + 1
    '
    Call PRINSTIKER(NEV1$, GRU1$, CLI1$, DEV1$, DGR1$, DCR1$, PLA1$, MAT1$, LLA1$, CAN1$, CRR1)
    '
    Call COMUNI("Se ha Procesado con Éxito\la Recepción Definitiva.")
    Unload Me
    '
End Sub

Sub PRINSTIKER(NEV1$, GRU1$, CLI1$, DEV1$, DGR1$, DCR1$, PLA1$, MAT1$, LLA1$, CAN1$, CRR1)
'
   '
   STIKAP$ = TRIM$(UCase$(Control$("", "STIKAPRO")))
   If Left$(STIKAP$, 4) <> "STIK" Then
     If Left$(STIKAP$, 5) <> "STICK" Then
       Exit Sub
     End If
   End If
   '
Printer.ScaleMode = 6
Printer.PaperSize = 256
Printer.Height = 45
Printer.Width = 91
Printer.Font = "Courier"
Printer.FontSize = 10
'
With Printer
    '
    XX$ = AJUSTI(NEV1$, 16) + AJUSTI(GRU1$, 10) + Mid$(CLI1$, 1, 15)
    '
    Printer.Print XX$
    '
    XX$ = ""
    Printer.Print XX$
    '
    XX$ = DEV1$
    Printer.Print XX$
    '
    XX$ = Mid(DGR1$, 1, 42)
    Printer.Print XX$
    '
    XX$ = Space(35) + "Plano"
    Printer.Print XX$
    '
    XX$ = AJUSTD(CStr(CRR1), 2) + Space(1) + CAN1$ + Space(1) + Mid$(DCR1$, 1, 20) + Space(1) + Mid$(PLA1$, 1, 12)
    Printer.Print XX$
    '
    XX$ = ""
    Printer.Print XX$
    '
    XX$ = "Material: " + AJUSTI(MAT1$, 16) + " Ll.: " + LLA1$
    Printer.Print XX$
    '
    .EndDoc
End With

'
'    FORIPRE$ = CONTROL$("", "STIKAPRO")
'    If TRIM$(FORIPRE$) <> "" Then
'      LU$ = LUDAT$
'      If EXISTE%(LU$ + FORIPRE$ + ".FRM") > 0 Then
'        '
'        TIPODOC$ = "'Stiker de Aprobación"
'        HOII% = HOY(HOS$)
'        FECHDOC$ = HOS$
'        NUMEDOC$ = CRR1
'          NBoRE& = Val(Label9)
'        DOCUORIG$ = "BR." + TRIM$(Str$(NBoRE&))
'        '
'        DESTIDOC% = 0
'        CACOLTAB1% = 0
'        CAITAB1% = 0
'        '
'        CODPARAM$(1) = "TRANSPOR"
'        CAPARDOC% = 11
'        '
'        Call PRINTDOC("STIKAPRO")   ' Remito de Facturacion
'        '
'      End If
'    End If
'    '
End Sub

Private Sub Form_Load()
    If Control("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture15.Visible = True
      Me.BackColor = &HFCD7B6
      Me.Frame1.BackColor = &HFCD7B6
      Me.Frame2.BackColor = &HFCD7B6
      Me.Frame3.BackColor = &HFCD7B6
      Me.Frame4.BackColor = &HFCD7B6
      Me.Frame5.BackColor = &HFCD7B6
      Me.Frame6.BackColor = &HFCD7B6
      Me.Frame7.BackColor = &HFCD7B6
      MARCOBARRA.Top = 5800
      Set MARCOBARRA.Container = Me.Picture15
    End If

End Sub

Private Sub salir_Click()
   Call Command1_Click
End Sub

'
Private Sub Text2_gotfocus()
    '
    BOTEXIT.SetFocus
    '
End Sub
Private Sub Text1_GotFocus()
    '
    BOTEXIT.SetFocus
    '
End Sub
Private Sub Text4_gotfocus()
    '
    BOTEXIT.SetFocus
    '
End Sub
Private Sub Text5_gotfocus()
    '
    BOTEXIT.SetFocus
    '
End Sub



