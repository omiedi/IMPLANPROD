VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form APROCLIB07 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aprobación de Material de Texto Libre"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   8460
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture999 
      Height          =   330
      Index           =   0
      Left            =   1470
      ScaleHeight     =   270
      ScaleWidth      =   1110
      TabIndex        =   31
      Top             =   525
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00CDDADA&
      Height          =   735
      Left            =   4905
      TabIndex        =   27
      Top             =   4650
      Width           =   2535
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   860
         TabIndex        =   29
         Top             =   240
         Width           =   1485
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Remito"
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
         Left            =   -1200
         TabIndex        =   28
         Top             =   310
         Width           =   1920
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00CDDADA&
      Height          =   735
      Left            =   225
      TabIndex        =   22
      Top             =   4650
      Width           =   4575
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo"
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
         Left            =   2760
         TabIndex        =   26
         Top             =   315
         Width           =   600
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   3450
         TabIndex        =   25
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "E.V."
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
         Left            =   -1290
         TabIndex        =   24
         Top             =   310
         Width           =   1920
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   720
         TabIndex        =   23
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.PictureBox Picture11 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   6960
      Left            =   7560
      ScaleHeight     =   6930
      ScaleWidth      =   1035
      TabIndex        =   16
      Top             =   -105
      Width           =   1065
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   120
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   35
         Top             =   6120
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   36
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton BOTEXIT 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   120
         Picture         =   "APROCLIB07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Rechazo del Material."
         Top             =   4440
         Width           =   650
      End
      Begin VB.CommandButton Command2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   120
         Picture         =   "APROCLIB07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Aprueba la Recepcion Provisoria."
         Top             =   5070
         Width           =   650
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   120
         Picture         =   "APROCLIB07.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Cancela - Cierra la Aplicación"
         Top             =   420
         Width           =   645
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00CDDADA&
      Height          =   855
      Left            =   3345
      TabIndex        =   13
      Top             =   5475
      Width           =   4095
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   1745
         TabIndex        =   21
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Deposito Destino"
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
         Left            =   -240
         TabIndex        =   20
         Top             =   420
         Width           =   1920
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00CDDADA&
      Height          =   855
      Left            =   225
      TabIndex        =   12
      Top             =   5475
      Width           =   3015
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad Recibida"
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
         Left            =   -110
         TabIndex        =   15
         Top             =   420
         Width           =   1920
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   1920
         TabIndex        =   14
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00CDDADA&
      Height          =   3255
      Left            =   225
      TabIndex        =   4
      Top             =   1305
      Width           =   7215
      Begin VB.TextBox Text2 
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
         Height          =   2415
         Left            =   180
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   630
         Width           =   6855
      End
      Begin VB.TextBox Text1 
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
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Top             =   240
         Width           =   5835
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Material"
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
         Left            =   90
         TabIndex        =   11
         Top             =   330
         Width           =   960
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CDDADA&
      Height          =   1005
      Left            =   225
      TabIndex        =   0
      Top             =   210
      Width           =   7215
      Begin VB.Label Label14 
         BackColor       =   &H00C0C0FF&
         Height          =   375
         Left            =   4320
         TabIndex        =   19
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFC0&
         Height          =   375
         Left            =   3480
         TabIndex        =   18
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label9 
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
         Height          =   255
         Left            =   2520
         TabIndex        =   17
         Top             =   360
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1200
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nº O/C"
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
         Left            =   120
         TabIndex        =   9
         Top             =   330
         Width           =   960
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   5955
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
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
         Left            =   4920
         TabIndex        =   7
         Top             =   330
         Width           =   960
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Top             =   615
         Width           =   975
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Proveedor"
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
         Left            =   120
         TabIndex        =   2
         Top             =   690
         Width           =   960
      End
      Begin VB.Label Label11 
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
         Height          =   285
         Left            =   2400
         TabIndex        =   1
         Top             =   615
         Width           =   4650
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "APROCLIB07.frx":075E
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "APROCLIB07.frx":0992
      TabIndex        =   32
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "APROCLIB07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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
      NBORE& = Val(Label9)
      CARECI# = Val(Label7)
      SQLX$ = " NumeBoRe = " & NBORE& & " "
      .FindFirst SQLX$
      If .NoMatch = False Then
        .Edit
          !CanToApro = 0
          !CanToRech = CARECI#
          !Status = 1 'APROBADA
          !USUAPRO = USERNAME
          !FechLibe = Date
          !USUAPRO = USERNAME   'NUEVO
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
        With OCOMDETA
          .FindFirst ABUS$
          If .NoMatch = False Then
            .Edit
            !Cant_Rec = !Cant_Rec - RECHFAL
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
    Call COMUNI("Se ha Procesado el Rechazo de la B.R. " + TRIM$(Str$(NBORE)))
    Unload Me
    '
99 BOTEXIT.Enabled = True
End Sub

Sub REMRECHLIB() ' (NPROVE%, DEPX%, REMIRECH$)
   '
   FORIPRE$ = TRIM$(CONTROL$("", "REMIRECH"))
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
   NC% = Abs(Val(LABEL5))
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
    NBORE& = Val(Label9)
    Call ABREORTRA
    With ListaMat
        SQLX$ = " Nro_Ocom LIKE '" & NUMEDO$ & "' "
        .FindFirst SQLX$
        If .NoMatch = False Then
          '
          'FORMATEO EL NUMERO DE BR
          n1$ = CStr(NBORE&)
          Do While Len(n1$) < 6
            n1$ = "0" + n1$
          Loop
          NBR2$ = "BR-" + n1$ + "-00"
          '
          NEV1$ = !Nro_Otrab
          GRU1$ = !Nro_Grup
          DCR1$ = !Descri_Cor
          PLA1$ = !NPLANO
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
          CLI1$ = rasocli$(NC%)
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
        SQLX$ = " NumeBoRe = " & NBORE& & " AND NumRem_Prov = '" & NUREPRO$ & "' "
        .FindFirst SQLX$
        If .NoMatch = False Then
          .Edit
            !CanToApro = CARECI#
            !CanToRech = 0
            !Status = 1 'APROBADA
            !USUAPRO = USERNAME
            !Status = 1
            !FechLibe = Date
            !USUAPRO = USERNAME   'NUEVO
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
   STIKAP$ = TRIM$(UCase$(CONTROL$("", "STIKAPRO")))
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
  Call APLICACIONSKINS(Me, Picture11)

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



