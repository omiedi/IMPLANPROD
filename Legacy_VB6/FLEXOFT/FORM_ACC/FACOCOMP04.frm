VERSION 5.00
Begin VB.Form FACOCOMP04 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facturación de Proveedores - Opciones"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5040
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture8 
      BackColor       =   &H80000016&
      Height          =   645
      Left            =   525
      ScaleHeight     =   585
      ScaleWidth      =   2370
      TabIndex        =   19
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
         Picture         =   "FACOCOMP04.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   20
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
         TabIndex        =   21
         Top             =   75
         Width           =   1785
      End
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00404000&
      Height          =   4740
      Left            =   4095
      ScaleHeight     =   4680
      ScaleWidth      =   915
      TabIndex        =   14
      Top             =   0
      Width           =   975
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   140
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   17
         Top             =   3045
         Width           =   640
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   18
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   140
         Picture         =   "FACOCOMP04.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   890
         Width           =   650
      End
      Begin VB.CommandButton Command2 
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
         Height          =   690
         Left            =   140
         Picture         =   "FACOCOMP04.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   200
         Width           =   650
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -285
         Picture         =   "FACOCOMP04.frx":075E
         Top             =   3360
         Width           =   1515
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Tablas de Validación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Index           =   1
      Left            =   210
      TabIndex        =   4
      Top             =   2520
      Width           =   3690
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   630
         ScaleHeight     =   585
         ScaleWidth      =   2370
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
            Height          =   600
            Left            =   0
            Picture         =   "FACOCOMP04.frx":2680
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Acceso Directo a Tabla de Imputaciones Contables de Compras"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Plan de Cuen- tas Contables"
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
            Left            =   945
            TabIndex        =   7
            ToolTipText     =   "Acceso a la Actualización del Plan de Cuentas Contables"
            Top             =   105
            Width           =   1425
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Operaciones Especiales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2325
      Index           =   0
      Left            =   210
      TabIndex        =   0
      Top             =   105
      Width           =   3690
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   840
         ScaleHeight     =   585
         ScaleWidth      =   2370
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
            Height          =   600
            Left            =   0
            Picture         =   "FACOCOMP04.frx":27CA
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
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   840
            TabIndex        =   13
            ToolTipText     =   "Click Aqui Importa Archivo GETRAMO"
            Top             =   80
            Width           =   1530
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   630
         ScaleHeight     =   585
         ScaleWidth      =   2370
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
            Height          =   600
            Left            =   0
            Picture         =   "FACOCOMP04.frx":2914
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
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   840
            TabIndex        =   10
            ToolTipText     =   "Click Aqui Importa Archivo GETRAMO"
            Top             =   80
            Width           =   1110
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   420
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   1
         Top             =   420
         Width           =   2430
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
            Height          =   600
            Left            =   0
            Picture         =   "FACOCOMP04.frx":2A5E
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
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   840
            TabIndex        =   3
            ToolTipText     =   "Click Aqui Importa Archivo GETRAMO"
            Top             =   75
            Width           =   1530
         End
      End
   End
End
Attribute VB_Name = "FACOCOMP04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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
    
    NC% = Abs(CLIPROACTIVO)
    If NC% < 1 Then
10    Call SELECHO("ACREDOR")
      NC% = XVALO(VALACT1$("ACREDOR"))
      If NC% < 1 Then
          GoTo 999
      End If
    End If
    '
20  If TICUEPRO%((-1) * NC%) <> 0 Then
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
    FIMPUP04.Label1 = Str$(NC%)
    FIMPUP04.Label2 = RASOCLI$((-1) * NC%)
    FIMPUP04.Label3 = DOMICLI$((-1) * NC%)
    FIMPUP04.Label6 = CPOSCLI$((-1) * NC%)
    FIMPUP04.Label12 = LOCACLI$((-1) * NC%)
    '
    FIMPUP04.List2.Clear
    FIMPUP04.Refresh
    '
    Call BLANARCH("!SELDOCAP")
    Call FRESHECHO("!SELDOCAP")
    '
' AQUI ESCRIBIR UNA RUTINA PARA SELECCIONAR DE SALDACRE LOS DOCUMENTOS A IMPUTAR Y CARGARLOS EN ARCHIVO 'SELDOCAP'
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM SACREPEN "
    SQLX$ = SQLX$ + "WHERE NuProve = " & NC% & " "
    SQLX$ = SQLX$ + "ORDER BY FeVencim ASC;"
    Set FacPenTemp = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With FacPenTemp
      On Error Resume Next
      .MoveLast
      If Err < 1 Then
        .MoveFirst
        Do While Not .EOF
          TIX$ = UCase$(Left$(!CodiCom_Lar, 2))
          SG% = (-1)
          IMPO# = (-1) * !IBruComp * SG%
          '
          SALDO# = !SaldAcre - !SaldDebe
          MEMI1% = 0: MONEDOCU% = 0: SALDOLA# = 0
          On Error Resume Next
          MEMI1% = !Mone_Emis Mod 6
          MONEDOCU% = MEMI1%
          SALDOLA# = !SaldAcre_Dola - !SaldDebe_Dola
          On Error GoTo 0
          '
          NROX$ = !CodiCom_Lar
          FEMI = CVINT(!FechEmisi)
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
               FIMPUP04.List2.AddItem Z$
               '
             ElseIf SALDO# <= (-0.005) Then
               '
               Z$ = Space$(60) + MKS$(0)
               Call REPLA(Z$, NROX$, 1, 18)
               Call REPLA(Z$, FECHATEX$(FEMI), 19, 8)
               Call REPLA(Z$, IMPTEX$, 27, 14)
               Call REPLA(Z$, SALTEX$, 41, 14)
               '
               Call REGAPP("!SELDOCAP", Z$)
               KQ& = KQ& + 1
               '
          End If
        .MoveNext
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
    With FIMPUP04
      .Label27 = "Saldo no Imputado:"
      .Label16 = ""
      .Label16.Visible = False
      .Label31.Visible = False
      .Label30.Visible = False
      .Label28.Visible = False
      '
      NROXEL$ = TRIM$(Left$(SPDX$, 18))
      .Label25 = Left$(SPDX$, 18)
      .Label24 = Mid$(SPDX$, 19, 8)
      .Label23 = Mid$(SPDX$, 27, 14)
      .Label22 = Mid$(SPDX$, 41, 14)
      .Label14 = Mid$(SPDX$, 41, 14)
      .Label18 = ""
      .Label10 = ""
      .Label20 = .Label14
      '
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
         .Label16 = FORMATNUM$(TICAM#, "F9.4")
         .Label28 = SIMBOLPES$(MEMI%)
         .Label32 = TRIM$(Str$(MEMI%))
         .Label16.Visible = True
         .Label28.Visible = True
         .Label30.Visible = True
         .Label31.Visible = True
         '
         If MEMI% > 1 Then
            ' VOLVER A CARGAR EL LIST2
            .List2.Clear
            With FacPenTemp
               On Error Resume Next
               .MoveLast
               If Err < 1 Then
                  On Error GoTo 0
                  .MoveFirst
                  Do While Not .EOF
                     TIX$ = UCase$(Left$(!CodiCom_Lar, 2))
                     IMPO# = !IBruComp
                     SALDO# = 0: MONEDOCU% = 1
                     On Error Resume Next
                     SALDO# = !SaldAcre_Dola - !SaldDebe_Dola
                     MONEDOCU% = !Mone_Emis
                     On Error GoTo 0
                     If MEMI% > 1 Then
                       If MEMI% <> MONEDOCU% Then GoTo 95
                     End If
                     NROX$ = !CodiCom_Lar
                     FEMI = CVINT(!FechEmisi)
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
                        FIMPUP04.List2.AddItem Z$
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
80  FIMPUP04.Show 1
    If CLIPROACTIVO = 0 Then
        GoTo 10
      Else
        GoTo 20
    End If
    '
999 Command3.Enabled = True
End Sub

Sub Command4_Click()
    Command4.Enabled = False
    Call OPORPA04.AJUSALP        ' Ajuste de Saldos de Proveedores
    Command4.Enabled = True
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
    FECHALU = Int(CDbl(CVDAT(FECHALIM%)))
    '
10  Call SELECHO("PROVED1")
    NC% = XVALO(VALACT1$("PROVED1"))
    If NC% < 1 Then Exit Sub
    '
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE (TipoMovim = 'FCOM' OR TipoMovim = 'AJUP') "
    SQLX$ = SQLX$ + "AND NuProve = " & NC% & " "
    SQLX$ = SQLX$ + "AND INT(FechEmisi) > " & FECHALU & " "
    SQLX$ = SQLX$ + "AND INT(FechConta) > " & FECHALU & " "
    SQLX$ = SQLX$ + "Order by CODICOM_LAR ASC "
    '
    Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    JJ& = 0
    With CABATEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        On Error GoTo 0
        Do While Not .EOF
          NCLIE% = !Nuprove
          IMPO# = !ValAbsComp
          If IMPO# >= 0.005 Then
            IMPOS$ = NUMSTRI$(IMPO#, 14, 2, 1)
            NPX& = XVALO(Mid$(!CodiCom_Lar, 13, 6))
            If Len(!CodiCom_Lar) < 18 Then
              NPX& = XVALO(Right$(!CodiCom_Lar, 6))
            End If
            If NPX& < 1 Then NPX& = 1
            Y$ = MKS$(NPX&) + AJUSTI$(!CodiCom_Lar, 18) + IMPOS$
            JJ& = JJ& + 1
            Call GRAREG("!SELCUECO", Y$, JJ&)
          End If
        .MoveNext
        Loop
      End If
    End With
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

    Call APERBASDAT("CABAN")
    '
    '\\\OT-06-0563-RG-21/12/06///
    RCLI$ = RASOCLI$((-1) * NC%)
    OPPX% = COMALTER%("Confirma la Anulación del Comprobante '" + TRIM$(DOCULA$) + "'\de " + RCLI$ + " \\No, Conservar\Si, Anular")
    If OPPX% <> 2 Then GoTo 10
    '\\\OT-06-0563-RG-FIN///
    '
    With CajaBanc
      .FindFirst "CodiCom_Lar = '" & DOCULA$ & "' AND NuProve = " & NC% & " "
20     If .NoMatch = False Then
        .Delete
        .FindNext "CodiCom_Lar = '" & DOCULA$ & "' AND NuProve = " & NC% & ""
        GoTo 20
      End If
    End With
    '
    With SacrePen
      .FindFirst "CodiCom_Lar = '" & DOCULA$ & "' AND NuProve = " & NC% & " "
25     If .NoMatch = False Then
        .Delete
        .FindNext "CodiCom_Lar = '" & DOCULA$ & "'AND NuProve = " & NC% & ""
        GoTo 25
      End If
    End With
    '
    ' RECONSTRUIR COMPOSICION DE DEUDA
    NumeClie% = 0
    NumeProv% = NC%
    SQLX$ = "SELECT * FROM ApliPago "
    SQLX$ = SQLX$ + "WHERE NuProve = " & NC% & " "
    Set ApliPag1 = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With ApliPag1
      .FindFirst "CoDocOr_Lar = '" & DOCULA$ & "' AND NuProve = " & NC% & " "
35     If .NoMatch = False Then
        CoComCor$ = !CoDocAp_Cor
        CoComLar$ = !CoDocAp_Lar
        FechEmi = Now
        IBruTot# = !ItotOrig
        IIvaTot# = !IIVAORIG
        NuorVen% = 0
        FeVenc = Now
        ImVenc# = (-1) * !ImpoApli
        '
        CajaBanc.FindFirst "CodiCom_Lar = '" & CoComLar$ & "' AND NuProve = " & NC% & " "
        If CajaBanc.NoMatch = False Then
          FechEmi = CajaBanc!FechEmisi
          FeVenc = CajaBanc!FechEmisi
        End If
        '
        Call GRACREPEN
        .FindNext "CoDocOr_Lar = '" & DOCULA$ & "'AND NuProve = " & NC% & ""
        GoTo 35
      End If
      '
      ' lo que sigue es identico para el caso de anular una nota de credito
      .FindFirst "CoDocAp_Lar = '" & DOCULA$ & "' AND NuProve = " & NC% & " "
45    If .NoMatch = False Then
        CoComCor$ = !CoDocOr_Cor
        CoComLar$ = !CoDocOr_Lar
        FechEmi = Now
        IBruTot# = !ItotOrig
        IIvaTot# = !IIVAORIG
        NuorVen% = 0
        FeVenc = Now
        ImVenc# = !ImpoApli
        '
        CajaBanc.FindFirst "CodiCom_Lar = '" & CoComLar$ & "' AND NuProve = " & NC% & " "
        If CajaBanc.NoMatch = False Then
          FechEmi = CajaBanc!FechEmisi
          FeVenc = CajaBanc!FechEmisi
        End If
        '
        Call GRACREPEN
        .FindNext "CoDocAp_Lar = '" & DOCULA$ & "'AND NuProve = " & NC% & ""
        GoTo 45
      End If
    End With
    '
99  JJ& = 0
    For J& = 1 To ULTREG("DETFACOM")
       XX$ = REGLEIDO("DETFACOM", J&)
       NP1% = CVI(Mid$(XX$, 1, 2))
       NF1$ = TRIM$(Mid$(XX$, 3, 18))
       If NF1$ <> DOCULA$ Then ' <> COMPROBANTE
         JJ& = JJ& + 1
         Call GRAREG("DETFACOM", XX$, JJ&)
       ElseIf NP1% <> NC% Then ' MISMO COMPROBANTE Y <> PROVEEDOR
         JJ& = JJ& + 1
         Call GRAREG("DETFACOM", XX$, JJ&)
       End If
     Next J&
     Call SETULTREG("DETFACOM", JJ&)
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
    RECUEP04.Text1 = TRIM$(Str$(NC%))
    RECUEP04.Show 1
    '
    GoTo 10
    '
End Sub
