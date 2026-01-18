VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FACOCOMP07 
   BackColor       =   &H00C7D9CC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facturación de Proveedores - Opciones"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4890
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture5 
      BackColor       =   &H0050765A&
      Height          =   4740
      Left            =   3960
      ScaleHeight     =   4680
      ScaleWidth      =   915
      TabIndex        =   20
      Top             =   0
      Width           =   975
      Begin VB.CommandButton Command2 
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
         Height          =   690
         Left            =   140
         Picture         =   "FACOCOMP07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   210
         Width           =   650
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Help"
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
         Left            =   140
         Picture         =   "FACOCOMP07.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   920
         Width           =   650
      End
      Begin VB.CommandButton Command5 
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
         Left            =   140
         Picture         =   "FACOCOMP07.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Visualizar Recepciones Valorizadas"
         Top             =   1620
         Width           =   650
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   17
      Top             =   1320
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture8 
      BackColor       =   &H80000016&
      Height          =   645
      Left            =   525
      ScaleHeight     =   585
      ScaleWidth      =   2370
      TabIndex        =   14
      Top             =   5250
      Width           =   2430
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
         Height          =   600
         Left            =   0
         Picture         =   "FACOCOMP07.frx":059E
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Edicion de Documentos de Compras"
         Top             =   0
         Width           =   560
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Consulta por Núme- ro de Documento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   630
         TabIndex        =   16
         Top             =   75
         Width           =   1785
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C7D9CC&
      Caption         =   "TABLAS DE VALIDACIÓN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   3690
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   630
         ScaleHeight     =   615
         ScaleWidth      =   2400
         TabIndex        =   5
         Top             =   420
         Width           =   2430
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
            Height          =   630
            Left            =   0
            Picture         =   "FACOCOMP07.frx":08A8
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Acceso Directo a Tabla de Imputaciones Contables de Compras"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Plan de Cuentas Contables"
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
            Left            =   700
            TabIndex        =   7
            ToolTipText     =   "Acceso a la Actualización del Plan de Cuentas Contables"
            Top             =   100
            Width           =   1425
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C7D9CC&
      Caption         =   "OPERACIONES ESPECIALES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2325
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   105
      Width           =   3690
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   840
         ScaleHeight     =   615
         ScaleWidth      =   2400
         TabIndex        =   11
         Top             =   1470
         Width           =   2430
         Begin VB.CommandButton Command4 
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
            Left            =   0
            Picture         =   "FACOCOMP07.frx":09F2
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Ajuste Directo de Saldos Acreedores"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Ajuste de Sal- dos Acreedores"
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
            Left            =   700
            TabIndex        =   13
            ToolTipText     =   "Click Aqui Importa Archivo GETRAMO"
            Top             =   100
            Width           =   1530
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   630
         ScaleHeight     =   615
         ScaleWidth      =   2400
         TabIndex        =   8
         Top             =   945
         Width           =   2430
         Begin VB.CommandButton Command3 
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
            Left            =   0
            Picture         =   "FACOCOMP07.frx":0B3C
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Imputación de Créditos y Anticipos de Proveedores"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Imputación de Creditos"
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
            Left            =   700
            TabIndex        =   10
            ToolTipText     =   "Click Aqui Importa Archivo GETRAMO"
            Top             =   50
            Width           =   1110
         End
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   420
         ScaleHeight     =   615
         ScaleWidth      =   2400
         TabIndex        =   1
         Top             =   420
         Width           =   2430
         Begin VB.CommandButton Command6 
            Caption         =   "?"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   1850
            Picture         =   "FACOCOMP07.frx":0C86
            TabIndex        =   19
            ToolTipText     =   "Aplicar Gasto Desaplicados a Factura"
            Top             =   0
            Visible         =   0   'False
            Width           =   560
         End
         Begin VB.CommandButton Command25 
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
            Left            =   0
            Picture         =   "FACOCOMP07.frx":0F90
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Anulación de Facturas, Débitos o Créditos de Proveedores"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Anulación de Documentos"
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
            Left            =   700
            TabIndex        =   3
            ToolTipText     =   "Click Aqui Importa Archivo GETRAMO"
            Top             =   50
            Width           =   1530
         End
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   750
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "FACOCOMP07.frx":129A
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   1590
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3600
      OleObjectBlob   =   "FACOCOMP07.frx":1314
      Top             =   1200
   End
End
Attribute VB_Name = "FACOCOMP07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub VALIFACTUGASTO(NC%, DOCULA$)
    '
    'VERIFICO SI TIENE FACTURAS EN LA TABLA FACTURAYGASTOS
    'SI TIENE FACTURAS ASOCIADAS PRESENTA LA LISTA
    CONDI$ = "Nume_Cli = " & NC% & " AND FACTURAPROV = '" & DOCULA$ & "' "
    If CantRegistros("FACTURAYGASTOS", CONDI$, "NOMESS") > 0 Then
            CONDI$ = "Nume_Cli = " & NC% & " AND FACTURAPROV = '" & DOCULA$ & "'  "
            TIPFAC$ = VALOBADA("FACTURAYGASTOS", "TIPO_FACTURA", CONDI$, "NOMESS")
            FACTASOCIADA$ = VALOBADA("FACTURAYGASTOS", "FACTURAPROV_ASOC", CONDI$, "NOMESS")
            If TIPFAC$ = "G" Then
               OPX% = COMALTER%("Factura Seleccionada Informada Como Gasto de Factura: " & FACTASOCIADA$ & "\\Cancelar Anulación\Continuar Anulación")
               If OPX% = 2 Then
                  ANUFACTGAS.MODO$ = "BORRAR"
                  CONDI$ = "Nume_Cli = " & NC% & " AND FACTURAPROV = '" & DOCULA$ & "'"
                  Call BORRAREGISTROS("FACTURAYGASTOS", CONDI$, REGAF&, "NOMESS")
               Else
                  ANUFACTGAS.MODO$ = "SALIR"
               End If
               Exit Sub
            End If

            OPX% = COMALTER%("Factura Seleccionada Posee Relación Sobre Otras Facturas\Deberá Anular Las Mismas\\Cancelar\Continuar ")
            If OPX% <> 2 Then
               Exit Sub
            Else
             'CARGO INFORMACION EN EL FRM DE DETALLE DE ANULACION
50           Call CARGA_ENCABEZADO(ANUFACTGAS.MSF1, "Proveedor/A24;Factura/A20;Imp.Fact./F20.2;Gastos Apl./F20.2;LIBRE/A0", ANUFACTGAS)
             ANUFACTGAS.TXTNFACTURA = DOCULA$
             ANUFACTGAS.TXTTIPOFACTURA = "MATERIALES"
             '
             CONDI$ = "Nume_Cli = " & NC% & " AND FACTURAPROV = '" & DOCULA$ & "'  "
             ANUFACTGAS.TXTIIMPORTE = FORMATNUM$(XVALO(VALOBADA("FACTURAYGASTOS", "IMPORTE_FACTURA", CONDI$, "NOMESS")), "F12.2")
             ANUFACTGAS.TXTNUMEPROVE = ""
             ANUFACTGAS.TXTNUMEPROVE = NC% ' AL CARGAR EL PROVEEDOR DISPARA EL EVENTO CHANGE Y MUESTRA HACE UN SHOW
             'Call ANUFACTGAS.Command1_Click
        End If
    End If

End Sub

Sub Command1_Click()
    Command1.Enabled = False
    Call CARCUEINGAS
    Command1.Enabled = True
End Sub

Private Sub Command2_Click()
    '
    Hide
    '
End Sub

Private Sub Command23_Click()
    '
    Command23.Enabled = False
    Call OPNOIN("Consulta por Número de Documento de Compra")
    'Call REPIORPA
    Command23.Enabled = True
    '
End Sub

Sub Command25_Click()
    Command25.Enabled = False
    Call ANUFACO
    Command25.Enabled = True
End Sub

Sub Command3_Click()
    '
    Command3.Enabled = False
    
    NICONFLIM% = NIVCONFI%
    NC% = Abs(CLIPROACTIVO)
    If NC% < 1 Then
10    Call SELECHO("ACREDOR")
      NC% = XVALO(VALACT1$("ACREDOR"))
      If NC% < 1 Then
          GoTo 999
      End If
    End If
    '
    If TICUEPRO%((-1) * NC%) <> 0 Then
      Call MENSERR(24, "Operacion Imposible. Cuenta de Gastos.")
      GoTo 10
    End If
    '
    MODOMONEI% = 0
    MODOMONE$ = Left$(TRIM$(UCase$(CONTROL$("GESACRED", "MODOMONE"))), 8)
    If MODOMONE$ = "MULTIMON" Then
       MODOMONEI% = 2
    End If
    '
    FIMPUP07.Label1 = Str$(NC%)
    FIMPUP07.Label2 = rasocli$((-1) * NC%)
    FIMPUP07.Label3 = DOMICLI$((-1) * NC%)
    FIMPUP07.Label6 = CPOSCLI$((-1) * NC%)
    FIMPUP07.Label12 = LOCACLI$((-1) * NC%)
    '
    FIMPUP07.List2.Clear
    FIMPUP07.Refresh
    '
    Call BLANARCH("!SELDOCAP")
    Call FRESHECHO("!SELDOCAP")
    '
    ' AQUI ESCRIBIR UNA RUTINA PARA SELECCIONAR DE SALDACRE LOS DOCUMENTOS A IMPUTAR Y CARGARLOS EN ARCHIVO 'SELDOCAP'
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM SACREPEN "
    SQLX$ = SQLX$ + "WHERE NuProve = " & NC% & " "
    SQLX$ = SQLX$ + "ORDER BY FeVencim ASC;"
    
    Dim FacPenTemp As ADODB.Recordset
    Set FacPenTemp = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    With FacPenTemp
    If Not (.EOF And .BOF) Then
        .MoveFirst
        Do While Not .EOF
          '
          NICON% = XVALO(!Niv_Clasi)
          If NICON% > NICONFLIM% Then GoTo 899     ' NO PERMITE REAPLICAR
          TIX$ = UCase$(Left$(SAFETEXT$(!CODICOM_LAR), 2))
          SG% = (-1)
          IMPO# = (-1) * XVALO(!IBruComp) * SG%
          SALDO# = XVALO(!SaldAcre) - XVALO(!SaldDebe)
          '
          MEMI1% = 0: MONEDOCU% = 0: SALDOLA# = 0
          On Error Resume Next
          MEMI1% = XVALO(!Mone_Emis) Mod 6
          MONEDOCU% = MEMI1%
          SALDOLA# = XVALO(!SaldAcre_DOLA) - XVALO(!SaldDebe_DOLA)
          On Error GoTo 0
          '
          NROX$ = SAFETEXT$(!CODICOM_LAR)
          FEMI = CVINT(!FECHEMISI)
          FEVE = CVINT(!FEVENCIM)
          '
          IMPTEX$ = CSTRING$(MKD$(IMPO#), 3, 12, 2, 2)
          SALTEX$ = CSTRING$(MKD$(Abs(SALDO#)), 3, 12, 2, 2)
          '
          If SALDO# >= 0.005 Then
               '
               SUMAFAC# = SUMAFAC# + IMPO#
               IMPO# = Abs(IMPO#)
               SUMADEU# = SUMADEU# + SALDO#
               '
               Z$ = Space$(128)
               Call REPLA(Z$, NROX$, 1, 18)
               Call REPLA(Z$, FECHATEX$(FEMI), 23, 8)
               Call REPLA(Z$, IMPTEX$, 32, 13)
               Call REPLA(Z$, FECHATEX$(FEVE), 47, 8)
               Call REPLA(Z$, SALTEX$, 57, 13)
               '
               If MODOMONEI% = 2 Then
                  Z$ = Left$(Z$ + Space$(160), 160)
                  Call REPLA(Z$, FORMATNUM$(MEMI1%, "I2"), 147, 2)
                  Call REPLA(Z$, FORMATNUM$(SALDOLA#, "F12.2"), 149, 12)
               End If
               '
               FIMPUP07.List2.AddItem Z$
               '
             ElseIf SALDO# <= (-0.005) Then
               '
               Z$ = Space$(60) + MKS$(0)
               Call REPLA(Z$, NROX$, 1, 18)
               Call REPLA(Z$, FECHATEX$(FEMI), 19, 8)
               Call REPLA(Z$, IMPTEX$, 27, 14)
               Call REPLA(Z$, SALTEX$, 41, 14)
               Call REGAPP("!SELDOCAP", Z$)
               KQ& = KQ& + 1
               '
          End If
899     .MoveNext
        Loop
      End If
    End With
    '
    If ULTREG&("!SELDOCAP") < 1 Then
      Call MENSERR(24, "No Hay Créditos Pendientes de Aplicación\en la Cuenta Elegida.")
      GoTo 10
    End If
    '
    SPDX$ = REGSEL$("!SELDOCAP")
    If Len(SPDX$) <= 4 Then GoTo 10
    '
    With FIMPUP07
      .Label25 = Left$(SPDX$, 18)
      .Label24 = Mid$(SPDX$, 19, 8)
      .Label23 = Mid$(SPDX$, 27, 14)
      .Label22 = Mid$(SPDX$, 41, 14)
      .Label14 = Mid$(SPDX$, 41, 14)
      .Label18 = ""
      .Label10 = ""
      .Label20 = .Label14
      
      NROXEL$ = TRIM$(Left$(SPDX$, 18))
      If MODOMONEI% = 2 Then
         NCX1% = XVALO(.Label1)
         FEMIX = FECHANUM(.Label24)
         MEMI% = MONEMISORP%(NROXEL$, NCX1%)
           If Abs(SALDACREPEN#(NROXEL$, NCX1%, 2)) < 0.005 Then
              MEMI% = 1    ' SI NO HAY SALDO EN DOLARES, TOMA SOLO LOS PESOS
           End If
         TICAM# = TICAMISORP(NROXEL$, FEMIX, NCX1%)
         TICAX1# = TICAM#: If MEMI% <= 1 Then TICAX1# = 1
         .Label23 = FORMATNUM$(ROUND#((XVALO(.Label23)) / TICAX1#), "F12.2")
         .Label22 = FORMATNUM$(ROUND#((XVALO(.Label22)) / TICAX1#), "F12.2")
         .Label14 = FORMATNUM$(ROUND#((XVALO(.Label14)) / TICAX1#), "F12.2")
         .Label20 = .Label14
         .Label27 = "Saldo N/I:"
         .Label28 = FORMATNUM$(TICAM#, "F9.4")
         .Label29 = SIMBOLPES$(MEMI%)
         .Label32 = TRIM$(Str$(MEMI%))
         .Label28.Visible = True
         .Label29.Visible = True
         .Label30.Visible = True
         .Label31.Visible = True
         .Label32.Visible = True
         '
         If MEMI% > 1 Then
            ' VOLVER A CARGAR EL LIST2
            .List2.Clear
            With FacPenTemp
               On Error Resume Next
               '.MoveLast
               If Err < 1 Then
                  On Error GoTo 0
                  .MoveFirst
                  Do While Not .EOF
                     TIX$ = UCase$(Left$(SAFETEXT$(CODICOM_LAR), 2))
                     IMPO# = XVALO(!IBruComp)
                     SALDO# = 0: MONEDOCU% = 1
                     On Error Resume Next
                     SALDO# = XVALO(!SaldAcre_DOLA) - XVALO(!SaldDebe_DOLA)
                     MONEDOCU% = XVALO(!Mone_Emis)
                     On Error GoTo 0
                     If MEMI% > 1 Then
                       If MEMI% <> MONEDOCU% Then GoTo 95
                     End If
                     NROX$ = SAFETEXT$(!CODICOM_LAR)
                     FEMI = CVINT(!FECHEMISI)
                     FEVE = CVINT(!FEVENCIM)
                     IMPTEX$ = CSTRING$(MKD$(IMPO#), 3, 12, 2, 2)
                     SALTEX$ = CSTRING$(MKD$(Abs(SALDO#)), 3, 12, 2, 2)
                     '
                     If SALDO# >= 0.005 Then
                        Z$ = Space$(128)
                        Call REPLA(Z$, NROX$, 1, 18)
                        Call REPLA(Z$, FECHATEX$(FEMI), 23, 8)
                        Call REPLA(Z$, IMPTEX$, 32, 13)
                        Call REPLA(Z$, FECHATEX$(FEVE), 47, 8)
                        Call REPLA(Z$, SALTEX$, 57, 13)
                        FIMPUP07.List2.AddItem Z$
                     End If
95                .MoveNext
                  Loop
               End If
            End With
            ' FIN DE VOLVER A CARGAR EL LIST2
         End If
      End If
    End With
    '
    FIMPUP07.Show 1
    If CLIPROACTIVO = 0 Then
      GoTo 10
    End If
    '
999 Command3.Enabled = True
End Sub

Sub Command4_Click()
    Command4.Enabled = False
    Call OPORPA07.AJUSALP        ' Ajuste de Saldos de Proveedores
    Command4.Enabled = True
End Sub


 Sub Command5_Click()
   Command5.Enabled = False
   '
10 Call SELECHO("ACREDOR")
   NP% = XVALO(TRIM$(VALACT1$("ACREDOR")))
   If NP% < 1 Then GoTo 99
   '3   Call SELECHO("!INDIOCOM/Indice General de Ordenes de Compra (Pendientes-Cumplidas-Anuladas)")

   CONDI$ = "NUMEBORE > 0 AND NumProv = " & NP% & " ORDER BY NUMEBORE"
   Call CARGALISEL("RECEPROV", "BOLRECEP", "NumeBoRe/F6;FechRecep/D8;NUMPROV/I5+PROVED2/A32", CONDI$, "NOMESS/DISTINCT")
20 Call SELECHO("RECEPROV/Recep. Proveedor: " & TRIM$(ECOARCH$("ACREDOR", MKI$(NP%))))
   NBR& = XVALO(TRIM$(VALACT1$("RECEPROV")))
   If NBR& < 1 Then GoTo 10
   '
   NBR1$ = CStr(NBR&)
   PROV$ = CStr(NP%)
   TOTAX# = 0
   J& = 0
   '
   SQLX$ = "SELECT SUM(BOLRECEP.CanToApro * MASTER.COSUNI) AS IMPORTE "
   SQLX$ = SQLX$ + " FROM BOLRECEP INNER JOIN MASTER ON BOLRECEP.COD_INTE = MASTER.CODINT"
   SQLX$ = SQLX$ + " WHERE BOLRECEP.NumeBoRe = " & NBR&
   'SQLX$ = SQLX$ + " AND OCOMDETA.Stat_Ocom < '9' AND OCOMDETA.Stat_Ocom >= '0'  "
   Set TOTAP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   With TOTAP
     Do While Not .EOF
        IMPO# = XVALO(!IMPORTE)
        TOTAX# = TOTAX# + IMPO#
        J& = J& + 1
     .MoveNext
     Loop
   End With
   TOTAP.Close
   Set TOTAP = Nothing
   '
   If J& < 1 Then Call COMUNI("No se Encontró El Detalle de La Recepción Seleccionada"): GoTo 10
   FILTX$ = TRIM$(NBR1$) & ";" & TRIM$(CStr(TOTAX#))
   ARCSQR$ = "REVE-LFE"
   Call STDFORM(ARCSQR$, FILTX$)
   '
   GoTo 20
   '
99 Command5.Enabled = True
   '
End Sub

Sub Command6_Click()
   Command6.Enabled = False
'   Call SELECHO("ACREDOR")
'   NCXI% = XVALO(TRIM$(VALACT1$("ACREDOR")))
'   If NCXI% > 0 Then
'     FACGASDES.TXTNUMEPROVE = NCXI%
'   End If
   OPX% = COMALTER%("Opciones de Trabajo \\Re-Aplicar o Desaplicar Facturas \Aplicación de Facturas de Gastos")
   If OPX% < 1 Or OPX% > 2 Then GoTo 99
   If OPX% = 1 Then
     FACGASDES.Caption = "Re-Aplicar o Desaplicar Facturas de Gastos"
       FACGASDES.Label2.Caption = "Boton derecho Para Des-Aplicar Factura de Gasto - Doble Click para Aplicar Gasto a Factura"
   Else
     FACGASDES.Caption = "Aplicación de Facturas de Gastos"
     FACGASDES.Label2.Caption = "Doble Click para Aplicar Gasto a Factura"
   End If
   FACGASDES.Show 1

99 Command6.Enabled = True
End Sub

Sub Command7_Click()
    Call HELPONLINE("FACOMP02")
End Sub

Sub ANUFACO()
 '
    If DERACCE%("ANUFACO", "Anulación de Facturas de Compra") < 2 Then
      Exit Sub
    End If
    '
    ' 1.- Ingresar Numero de Proveedor
    '     ****************************
    '
    FECHALIM% = XVALO(CONTROL$("", "CIEMECOM"))
    'jandy sql
    FECHALU = FETEXCOR1$(FECHALIM%)
    'FECHALU = Int(CDbl(CVDAT(FECHALIM%)))
    '
10  Call SELECHO("PROVED1")
    NC% = XVALO(VALACT1$("PROVED1"))
    If NC% < 1 Then Exit Sub
    '
    Call APERBASDAT("CABAN")
    Dim CABATEMP As ADODB.Recordset
    Set CABATEMP = New ADODB.Recordset
    CABATEMP.CursorLocation = adUseClient
    
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE (TipoMovim = 'FCOM' OR TipoMovim = 'AJUP') "
    SQLX$ = SQLX$ + "AND NuProve = " & NC% & " "
    'SQLX$ = SQLX$ + "AND INT(FechEmisi) > " & FECHALU & " "
    'SQLX$ = SQLX$ + "AND INT(FechConta) > " & FECHALU & " "
    SQLX$ = SQLX$ + "AND FechEmisi > '" & FECHALU & "' "
    SQLX$ = SQLX$ + "AND FechConta > '" & FECHALU & "' "
    SQLX$ = SQLX$ + "Order by CODICOM_LAR ASC "
    '
    'Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
25  On Error Resume Next
    CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
    End If
    On Error GoTo 0
    
    '
    JJ& = 0
    NIVCONXXX% = NIVCONFI%
    With CABATEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        On Error GoTo 0
        Do While Not .EOF
          '\\\OT-08-1032-FL-12/11/08 VERIFICA EL NIVEL DEL USUARIO PARA
          If XVALO(!Niv_Clasi) > NIVCONXXX% Then GoTo 70
          '\\\OT-08-1032-FL-FIN
       
          NCLIE% = !NUPROVE
          IMPO# = !ValAbsComp
          If IMPO# >= 0.005 Then
            IMPOS$ = NUMSTRI$(IMPO#, 14, 2, 1)
            NPX& = XVALO(Mid$(!CODICOM_LAR, 13, 6))
            If Len(!CODICOM_LAR) < 18 Then
              NPX& = XVALO(Right$(!CODICOM_LAR, 6))
            End If
            If NPX& < 1 Then NPX& = 1
            y$ = MKS$(NPX&) + AJUSTI$(!CODICOM_LAR, 18) + IMPOS$
            JJ& = JJ& + 1
            Call GRAREG("!SELCUECO", y$, JJ&)
          End If
70      .MoveNext
        Loop
      End If
    End With
    CABATEMP.Close
    Set CABATEMP = Nothing
    
    Call SETULTREG("!SELCUECO", JJ&)
    Screen.MousePointer = 1
    '
    If ULTREG&("!SELCUECO") < 1 Then
       Call MENSERR(24, "La Cuenta Elegida no Registra\Comprobantes de Compra que Puedan Anularse.")
       GoTo 10
    End If
    '
    Call CIERRARCH("!SELCUECO")
    Call FRESHECHO("!SELCUECO")
    '
    Call SELECHO("!SELCUECO")
    If VALACT1$("!SELCUECO") = "" Then GoTo 10
    REGCUEC& = XVALO(VALACT1$("!SELCUECO"))
    If REGCUEC& < 1 Then GoTo 10
    '
    DOCULA$ = Left$(VALACT2$("!SELCUECO"), 18)
    NUOPANU& = XVALO(Mid$(DOCULA$, 13, 6))
    
    '\\\OT-08-0220-FL-06/05/08///
    If VERICONCILIA%(DOCULA$) < 1 Then Exit Sub
    '\\\OT-08-0220-FL-FIN///
    Call APERBASDAT("CABAN")
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "AHP") > 0 Then
       Call VALIFACTUGASTO(NC%, DOCULA$)
       If ANUFACTGAS.MODO$ = "SALIR" Then
          Exit Sub
       ElseIf ANUFACTGAS.MODO$ = "BORRAR" Then
          GoTo 50 'SALTEO LA CONFIRMACION DE ANULACION POR QUE YA BORRO
       End If
    End If

    '\\\OT-06-0563-RG-21/12/06///
    RCLI$ = rasocli$((-1) * NC%)
    OPPX% = COMALTER%("Confirma la Anulación del Comprobante '" + TRIM$(DOCULA$) + "'\de " + RCLI$ + " \\No, Conservar\Si, Anular")
    If OPPX% <> 2 Then GoTo 10
    '\\\OT-06-0563-RG-FIN///
    '
    'JANDY SQL
50  strsql = "Delete CajaBanc where CodiCom_Lar = '" & DOCULA$ & "' "
    strsql = strsql & " AND NuProve = " & NC%
    
    strsql = strsql & " Delete SacrePen where CodiCom_Lar = '" & DOCULA$ & "'"
    strsql = strsql & " AND NuProve = " & NC%
    
    'se agrego borrado del registro en aplipago
    strsql = strsql & " Delete ApliPago where CoDocOr_Lar = '" & DOCULA$ & "'"
    strsql = strsql & " OR CoDocAp_Lar = '" & DOCULA$ & "'"
    strsql = strsql & " AND NuProve = " & NC%
    
    Call EJECUTACOMANDO(strsql)


'    With CajaBanc
'      .FindFirst "CodiCom_Lar = '" & DOCULA$ & "' AND NuProve = " & NC% & " "
'20     If .NoMatch = False Then
'        .Delete
'        .FindNext "CodiCom_Lar = '" & DOCULA$ & "' AND NuProve = " & NC% & ""
'        GoTo 20
'      End If
'    End With
'    '
'    With SacrePen
'      .FindFirst "CodiCom_Lar = '" & DOCULA$ & "' AND NuProve = " & NC% & " "
'25     If .NoMatch = False Then
'        .Delete
'        .FindNext "CodiCom_Lar = '" & DOCULA$ & "'AND NuProve = " & NC% & ""
'        GoTo 25
'      End If
'    End With
    '
    
    'JANDY SQL VER DE HACER UN JOIN ENTRE APLIPAGO Y CAJABANC
    
    ' RECONSTRUIR COMPOSICION DE DEUDA
    NumeClie = 0
    NumeProv% = NC%
    Dim Aplipag1 As ADODB.Recordset
    SQLX$ = "SELECT * FROM ApliPago "
    SQLX$ = SQLX$ + "WHERE NuProve = " & NC% & " "
    'Set Aplipag1 = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
     SQLX$ = SQLX$ + "AND  CoDocOr_Lar = '" & DOCULA$ & "'"
    Set Aplipag1 = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    With Aplipag1
    If Not (.EOF And .BOF) Then
      .MoveFirst
      Do While Not .EOF
'      .FindFirst "CoDocOr_Lar = '" & DOCULA$ & "' AND NuProve = " & NC% & " "
'35     If .NoMatch = False Then
        CoComCor$ = !CODOCAP_COR
        CoComLar$ = !codocap_lar
        FechEmi = Now
        IBruTot# = !ITOTORIG
        IIvaTot# = !IIVAORIG
        NuorVen% = 0
        FeVenc = Now
        ImVenc# = (-1) * !IMPOAPLI
        '
        'CajaBanc.FindFirst "CodiCom_Lar = '" & CoComLar$ & "' AND NuProve = " & NC% & " "
        'If CajaBanc.NoMatch = False Then
        Dim CajaBanc1 As ADODB.Recordset
        strsql = "Select * from CajaBanc where  CodiCom_Lar = '" & CoComLar$ & "'"
        strsql = strsql & "AND NuProve = " & NC%
        Set CajaBanc1 = FLEXCONN.Execute(FILTSQL$(strsql))
        If Not (CajaBanc1.EOF And CajaBanc1.BOF) Then
          FechEmi = CajaBanc1!FECHEMISI
          FeVenc = CajaBanc1!FECHEMISI
        End If
        CajaBanc1.Close
        Set CajaBanc1 = Nothing
        '
        Call GRACREPEN
       ' .FindNext "CoDocOr_Lar = '" & DOCULA$ & "'AND NuProve = " & NC% & ""
       ' GoTo 35
      .MoveNext
      Loop
    End If
    End With
    Aplipag1.Close
    Set Aplipag1 = Nothing
      
'Jandy sql genero una nueva consulta,extraigo los datos y grabo
      '
      ' lo que sigue es identico para el caso de anular una nota de credito
      
    Dim Aplipag2 As ADODB.Recordset
    SQLX$ = "SELECT * FROM ApliPago "
    SQLX$ = SQLX$ + "WHERE NuProve = " & NC% & " "
    SQLX$ = SQLX$ + "AND CoDocAp_Lar = '" & DOCULA$ & "'"
    Set Aplipag2 = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    With Aplipag2
    If Not (.EOF And .BOF) Then
      .MoveFirst
      Do While Not .EOF
'      .FindFirst "CoDocAp_Lar = '" & DOCULA$ & "' AND NuProve = " & NC% & " "
'45    If .NoMatch = False Then
        CoComCor$ = !CODOCOR_COR
        CoComLar$ = !codocor_lar
        FechEmi = Now
        IBruTot# = !ITOTORIG
        IIvaTot# = !IIVAORIG
        NuorVen% = 0
        FeVenc = Now
        ImVenc# = !IMPOAPLI
        '
        'CajaBanc.FindFirst "CodiCom_Lar = '" & CoComLar$ & "' AND NuProve = " & NC% & " "
        'If CajaBanc.NoMatch = False Then
        Dim CajaBanc2 As ADODB.Recordset
        strsql = "Select * from CajaBanc where  CodiCom_Lar = '" & CoComLar$ & "'"
        strsql = strsql & "AND NuProve = " & NC%
        Set CajaBanc2 = FLEXCONN.Execute(FILTSQL$(strsql))
        If Not (CajaBanc2.EOF And CajaBanc2.BOF) Then
          FechEmi = CajaBanc2!FECHEMISI
          FeVenc = CajaBanc2!FECHEMISI
        End If
        CajaBanc2.Close
        Set CajaBanc2 = Nothing
        '
        Call GRACREPEN
       ' .FindNext "CoDocAp_Lar = '" & DOCULA$ & "'AND NuProve = " & NC% & ""
       ' GoTo 45
       .MoveNext
       Loop
      End If
    End With
    Aplipag2.Close
    Set Aplipag2 = Nothing
    
    
    '
99  JJ& = 0

    CONDI$ = " Nume_Cli= " & NC%
    CONDI$ = CONDI$ + " AND FACTURAPROV = '" & DOCULA$ & "' "
    Call BORRAREGISTROS("DETFACOMSQL", CONDI$, REGAF&, "NOMESS")
    
'    For j& = 1 To ULTREG("DETFACOM")
'       XX$ = REGLEIDO("DETFACOM", j&)
'       NP1% = CVI(Mid$(XX$, 1, 2))
'       NF1$ = TRIM$(Mid$(XX$, 3, 18))
'       If NF1$ <> DOCULA$ Then ' <> COMPROBANTE
'         JJ& = JJ& + 1
'         Call GRAREG("DETFACOM", XX$, JJ&)
'       ElseIf NP1% <> NC% Then ' MISMO COMPROBANTE Y <> PROVEEDOR
'         JJ& = JJ& + 1
'         Call GRAREG("DETFACOM", XX$, JJ&)
'       End If
'     Next j&
'     Call SETULTREG("DETFACOM", JJ&)
     '
     JJ& = 0
     For J& = 1 To ULTREG("FACBONCO")
       XX$ = REGLEIDO("FACBONCO", J&)
       NP1% = CVI(Mid$(XX$, 1, 2))
       NF1$ = TRIM$(Mid$(XX$, 9, 18))
       If NF$ <> NF1$ Then ' <> COMPROBANTE
         JJ& = JJ& + 1
         Call GRAREG("FACBONCO", XX$, JJ&)
       ElseIf NC% <> NP1% Then ' MISMO COMPROBANTE Y <> PROVEEDOR
         JJ& = JJ& + 1
         Call GRAREG("FACBONCO", XX$, JJ&)
       End If
     Next J&
     Call SETULTREG("FACBONCO", JJ&)
     '
    Screen.MousePointer = 1
    Call COMUNI("ANULACIóN COMPLETA.\  \Revise Ahora el Resumen\de Cuenta Corriente,\Verificando la Correcta Composición de Deuda.")
    RECUEP07.Text1 = TRIM$(Str$(NC%))
    RECUEP07.Show 1
    '
    GoTo 10
    '
End Sub

Private Sub Form_Load()

    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "AHP") > 0 Then
       Command6.Visible = True
    End If
    '
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me, Picture5)
End Sub
