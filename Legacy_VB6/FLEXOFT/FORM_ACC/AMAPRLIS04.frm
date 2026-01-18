VERSION 5.00
Begin VB.Form AMAPRLIS04 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0E0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Maestro de Proveedores - Consultas y Listados"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   5235
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0C0&
      Caption         =   "Consultas - Listados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   160
      TabIndex        =   0
      Top             =   150
      Width           =   4005
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1155
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   14
         Top             =   1365
         Width           =   2325
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
            Picture         =   "AMAPRLIS04.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Listado de Números de C.A.I."
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Listado de CAI   de Proveedores"
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
            TabIndex        =   15
            Top             =   80
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   840
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   4
         Top             =   840
         Width           =   2325
         Begin VB.CommandButton Command5 
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
            Left            =   -20
            Picture         =   "AMAPRLIS04.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Listado de Cuentas Inactivas"
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
            TabIndex        =   6
            Top             =   80
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   525
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   7
         Top             =   315
         Width           =   2325
         Begin VB.CommandButton Command8 
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
            Left            =   -20
            Picture         =   "AMAPRLIS04.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Listado General   y Clasificado"
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
            Left            =   735
            TabIndex        =   9
            Top             =   50
            Width           =   1590
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0E0C0&
         Caption         =   "Salida"
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
         Left            =   420
         TabIndex        =   1
         Top             =   2050
         Width           =   3165
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0E0C0&
            Caption         =   "Impresión"
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
            Left            =   1785
            TabIndex        =   3
            Top             =   250
            Value           =   -1  'True
            Width           =   1275
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0E0C0&
            Caption         =   "Pantalla"
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
            Index           =   0
            Left            =   315
            TabIndex        =   2
            Top             =   250
            Width           =   1380
         End
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004000&
      Height          =   3585
      Left            =   4305
      ScaleHeight     =   3525
      ScaleWidth      =   900
      TabIndex        =   10
      Top             =   0
      Width           =   960
      Begin VB.CommandButton command1 
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
         Picture         =   "AMAPRLIS04.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Cierra - Abandona la Aplicación"
         Top             =   210
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
         Height          =   640
         Left            =   105
         Picture         =   "AMAPRLIS04.frx":0A68
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Libreta de Direcciones de Mails"
         Top             =   850
         Width           =   650
      End
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
         Height          =   640
         Left            =   105
         Picture         =   "AMAPRLIS04.frx":0D72
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Listado de Direcciones Postales"
         Top             =   1490
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -210
         Picture         =   "AMAPRLIS04.frx":107C
         Stretch         =   -1  'True
         Top             =   2470
         Width           =   1410
      End
   End
End
Attribute VB_Name = "AMAPRLIS04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Call CIERRARCH("*.*")
   Hide
End Sub

 Sub Command2_Click()
   '
   Command2.Enabled = False
   CMQ% = COMALTER%("Opciones Disponibles\\Listado Impreso\Exportar")
   If CMQ% < 1 Or CMQ% > 2 Then GoTo 99
   '
   Screen.MousePointer = 11
   'Consulta SQL
   Call APERBASDAT("PROVE")
   SQLX$ = " Select * from Proved12 "
   SQLX$ = SQLX$ + " ORDER BY Nume_Cli "
   Set PROVTEMP = Badaprov.OpenRecordset(SQLX$, dbOpenSnapshot)
     With PROVTEMP
       On Error Resume Next
       .MoveFirst
       'Valida que la tabla no este vacia
       If Err Then
         Call MENSERR(24, "Ningún Registro Seleccionado según Condición SQL")
         On Error GoTo 0
         Screen.MousePointer = 1
         GoTo 90
       End If
       On Error GoTo 0
     End With
   '
   If CMQ% = 1 Then         ' IMPRESION
    With PROVTEMP
      JJ& = 0
      Call COPYSTRU("LITELMAI", "LITELPRO")
      Call REGBLAN("LITELPRO")
      Do While .EOF = False
           NP% = !Nume_Cli
           NC% = (-1) * NP%
        EMA$ = EMAILCLI$(NC%)
        EMB$ = TEL1CLI$(NC%)
        EMC$ = FAXCLI$(NC%)
        If TRIM$(EMA$ + EMB$ + EMC$) <> "" Then
          'TTXX$ = String$(2, 0) + Space$(126)
          Call REPLA(TTXX$, MKI$(NP%), 1, 2)
          Call REPLA(TTXX$, EMB$, 33, 24)
          Call REPLA(TTXX$, EMC$, 57, 24)
          Call REPLA(TTXX$, EMA$, 81, 48)
          JJ& = JJ& + 1
          Call GRAREG("LITELPRO", TTXX$, JJ&)
        End If
        .MoveNext
        Loop
    End With
      '
      Call SETULTREG("LITELPRO", JJ&)
      Call CIERRARCH("LITELPRO")
      Call CIERRARCH("*.*")
      '
      Call FINAL("*LITELPRO")
      GoTo 99
      '
   End If
   '
   With PROVTEMP            'EXPORTACIÓN
     .MoveFirst
     JJ& = 0
     Call REGBLAN("!BCONTACT")
     Do While (.EOF = False)
      NC1% = (-1) * !Nume_Cli
        NC% = NC1%
        EMA$ = EMAILCLI$(NC%)
        If TRIM$(EMA$) <> "" Then
          If InStr(EMA$, "@") > 0 Then
             TTXX$ = String$(2, 0) + Space$(1022)
             RACLI$ = RASOCLI$(NC%)
             Call REPLA(TTXX$, RACLI$, 3, 78)
             Call REPLA(TTXX$, "Sres.", 81, 8)
             Call REPLA(TTXX$, RACLI$, 145, 64)
             Call REPLA(TTXX$, EMA$, 217, 64)
             Call REPLA(TTXX$, DOMICLI$(NC%), 768 + 1, 48)
             Call REPLA(TTXX$, CPOSCLI$(NC%), 768 + 49, 16)  ' cpost empresa
             Call REPLA(TTXX$, LOCACLI$(NC%), 768 + 65, 48)  ' localidad empresa
             Call REPLA(TTXX$, DEPROVI$(PROVCLI$(NC%)), 768 + 113, 32) ' provincia empresa
             Call REPLA(TTXX$, PAISCLI$(NC%), 768 + 145, 32) ' pais empresa
             Call REPLA(TTXX$, TEL1CLI$(NC%), 768 + 177, 40) ' telef.empresa
             Call REPLA(TTXX$, FAXCLI$(NC%), 768 + 217, 20) ' fax empresa
             JJ& = JJ& + 1
             Call GRAREG("!BCONTACT", TTXX$, JJ&)
          End If
        End If
      .MoveNext
     Loop
   End With
   '
   Call SETULTREG("!BCONTACT", JJ&)
   Call CIERRARCH("!BCONTACT")
   MSGXX$ = "Se ha Exportado con éxito en "
   MSGXX$ = MSGXX$ + LUCOM$ + "BCONTACT.WKF"
   'OT-05-0710-WAR-03/11/05///
   MsgBox (MSGXX$)
   'Call COMUNI(MSGXX$) se deja con msgbox xq la funcion
   'comuni no deja escribir en el texto el caracter \
   'OT-05-0710-WAR-FIN///
   Call CIERRARCH("*.*")
   '
99 Command2.Enabled = True
   Screen.MousePointer = 1
90  End Sub

Sub Command3_Click()
    Command3.Enabled = False
    Screen.MousePointer = 11
    '
    JJ& = 0
    FIN& = ULTREG&("PROVED1")
    '
    For U& = 1 To FIN&
        Call TRACE(20, U&, FIN&)
        '
        REPROVE$ = REGLEIDO$("PROVED1", U&)
        NPP% = CVI(Left$(REPROVE$, 2))
        '
        NUCAI$ = NUMECAI$(NPP%, FECHVENCAI%)
        X$ = REGBLAN("NUMECAI")
        Call REPLA(X$, MKI$(NPP%), 1, 2)
        Call REPLA(X$, NUCAI$, 3, 24)
        Call REPLA(X$, MKI$(FECHVENCAI%), 27, 2)
        JJ& = JJ& + 1
        Call GRAREG("NUMECAI", X$, JJ&)
    Next U&
    Call SETULTREG("NUMECAI", JJ&)
    '
    Call FINAL("*NUMECAI")
    Screen.MousePointer = 1
    Command3.Enabled = True

End Sub

 Sub Command4_Click()
Command4.Enabled = False
   '
Screen.MousePointer = 11
   '
Call APERBASDAT("PROVE")
'    FIN& = ULTREG&("DEUDOR1")
'    JJ& = 0
'    Call COPYSTRU("BADACLIE", "BADACLIX")

'Consulta sql
 SQLX$ = " SELECT * FROM PROVED12 "
'Filtro unicamente las cuentas corrientes
 SQLX$ = SQLX$ + " WHERE Tcpr_Cli <> 1 "
 SQLX$ = SQLX$ + " ORDER BY Raso_Cli ASC "
 Set PROVTEMP = Badaprov.OpenRecordset(SQLX$, dbOpenSnapshot)
     With PROVTEMP
       On Error Resume Next
       .MoveLast
       'Valida que la tabla no este vacia
       If Err Then
         Call MENSERR(24, "Ningún Registro Seleccionado según Condición SQL")
         GoTo 90
       End If
       registros& = .RecordCount
       FIN& = registros&
       .MoveFirst
    End With
With PROVTEMP
    JJ& = 0
    Call REGBLAN("BADACLIX")
 Do While (.EOF = False)
'    For U& = 1 To FIN&
'      Call TRACE(20, U&, FIN&)
'      X$ = REGLEIDO$("DEUDOR1", U&)
'      NC% = CVI(Left$(X$, 2))
       NP% = !Nume_Cli
       NC% = (-1) * NP%
     If NC% < 0 Then
       EMA$ = RASOCLI$(NC%)
       EMB$ = DOMICLI$(NC%)
       EMC$ = CPOSCLI$(NC%)
       EMD$ = LOCACLI$(NC%)
       EME$ = DEPROVI$(PROVCLI$(NC%))
       EMF$ = PAISCLI$(NC%)
       '
       If TRIM$(EMA$ + EMB$ + EMC$ + EMD$ + EME$ + EMF$) <> "" Then
         TTXX$ = String$(2, 0) + Space$(510)
         Call REPLA(TTXX$, MKI$(NC%), 1, 2)
         Call REPLA(TTXX$, EMA$, 3, 62)
         Call REPLA(TTXX$, EMB$, 97, 64)
         Call REPLA(TTXX$, EMC$, 161, 16)
         Call REPLA(TTXX$, EMD$, 177, 48)
         Call REPLA(TTXX$, EME$, 225, 32)
         Call REPLA(TTXX$, EMF$, 257, 24)
         JJ& = JJ& + 1
         Call GRAREG("BADACLIX", TTXX$, JJ&)
       End If
     End If
     .MoveNext
    Loop
'   Next U&
End With
   '
   Call SETULTREG("BADACLIX", JJ&)
   Call CIERRARCH("BADACLIX")
   Call CIERRARCH("*.*")
   '
   Call FINAL("*LIPOSTAL")
   Screen.MousePointer = 1
   '
   Command4.Enabled = True
90 End Sub
Sub Command5_Click()
Command5.Enabled = False
Call COMUNI("No Disponible por el Momento")
Command5.Enabled = True
Exit Sub
'Exit Sub
    Command5.Enabled = False
    Call SELECHO("SELFECHA")
    FECHLI$ = VALACT1$("SELFECHA")
    FLIM% = FECHANUM(FECHLI$)
    If FLIM% > 0 Then
       Screen.MousePointer = 11
       Dim ACTI%(32767)
       FIN& = ULTREG&("FACTUR")
       For U& = 1 To FIN&
         Call TRACE(20, U&, FIN&)
         X$ = REGLEIDO$("FACTUR", U&)
         NC% = CVI(Left$(X$, 2))
         If NC% > 0 Then
           If NC% <= 32767 Then
             If CVI(Mid$(X$, 7, 2)) > ACTI%(NC%) Then
               ACTI%(NC%) = CVI(Mid$(X$, 7, 2))
             End If
           End If
         End If
       Next U&
       '
       RELISA& = 0
       CX& = ULTREG&("DEUDOR1")
       For IL& = 1 To CX&
         Call TRACE(20, IL&, CX&)
         NC% = CVI(Left$(REGLEIDO$("DEUDOR1", IL&), 2))
         If ACTI%(NC%) < FLIM% Then
            Y$ = REGBLAN$("LISACLI")
            SSCC# = SALDCLI#(NC%)
            Call REPLA(Y$, MKI$(NC%), 1, 2)
            Call REPLA(Y$, MKD$(SSCC#), 3, 8)
            Call REPLA(Y$, MKD$(SSCC#), 11, 8)
            FEX = ACTI%(NC%)
            TTXX$ = "Inactiva Desde " + FECHATEX$(FEX)
            Call REPLA(Y$, TTXX$, 19, 30)
            '
            RELISA& = RELISA& + 1
            Call GRAREG("LISACLI", Y$, RELISA&)
         End If
       Next IL&
       '
       Call SETULTREG("LISACLI", RELISA&)
       Call CIERRARCH("*.*")
       Screen.MousePointer = 1
       Call CONECRUN("*DEUINAC", "Listado de Cuentas Inactivas")
    End If
    Command5.Enabled = True
End Sub

Sub Command8_Click()
    If Option1(0).Value = True Then
        Option2.Value = True
    End If
    Command8.Enabled = False
    FORLIPRO04.Show 1
    Command8.Enabled = True
End Sub

Sub LILIMICRE()
   '
   Screen.MousePointer = 11
   '
   Dim DEUCUECO#(16384), CHEPEAC#(16384)
   '
   FIN& = ULTREG&("SVD202")
   For IL& = 1 To FIN&
      Call TRACE(20, IL&, FIN&)
      SS$ = REGLEIDO$("SVD202", IL&)
      NCI% = CVI(Left$(SS$, 2))
      If NCI% > 0 Then
         If NCI% <= 16384 Then
            TI% = Asc(Mid$(SS$, 3, 1))
            If TI% > 7 Then TI% = 2
            SALDEU# = CVD(Mid$(SS$, 11, 8))
            If TI% > 3 Then SALDEU# = (-1) * Abs(SALDEU#)
            DEUCUECO#(NCI%) = DEUCUECO#(NCI%) + SALDEU#
            GoTo 19
         End If
      End If
19 Next IL&
   '
   HOII% = HOY(HOS$)
   FIN& = ULTREG&("CHECAR")
   For IL& = 1 To FIN&
      Call TRACE(20, IL&, FIN&)
      XX$ = REGLEIDO$("CHECAR", IL&)
      NCI% = CVI(Mid$(XX$, 3, 2))
      If NCI% > 0 Then
         If NCI% <= 16384 Then
            If CVI(Mid$(XX$, 41, 2)) >= HOII% Then
               CHEPEAC#(NCI%) = CHEPEAC#(NCI%) + CVD(Mid$(XX$, 33, 8))
            End If
         End If
      End If
   Next IL&
   '
   FIN& = ULTREG&("DEUDOR1")
   JJ& = 0
   For IL& = 1 To FIN&
      Call TRACE(20, IL&, FIN&)
      XX$ = REGLEIDO$("DEUDOR1", IL&)
      NCI% = CVI(Left$(XX$, 2))
      If NCI% > 0 Then
         If NCI% <= 16384 Then
            If Abs(DEUCUECO#(NCI%)) + Abs(CHEPEAC#(NCI%)) >= 0.005 Then
               YY$ = REGLEIDO$("DEUDOR2", IL&)
               CRELIM# = CVS(Mid$(YY$, 121, 4))
               DISPO# = CRELIM# - DEUCUECO#(NCI%) - CHEPEAC#(NCI%)
               EXCES# = 0
               If DISPO# < 0 Then
                  EXCES# = Abs(DISPO#)
                  DISPO# = 0
               End If
               ZZ$ = REGBLAN$("LILICRE")
               Call REPLA(ZZ$, MKI$(NCI%), 1, 2)
               Call REPLA(ZZ$, MKD$(CRELIM#), 3, 8)
               Call REPLA(ZZ$, MKD$(DEUCUECO#(NCI%)), 11, 8)
               Call REPLA(ZZ$, MKD$(CHEPEAC#(NCI%)), 19, 8)
               Call REPLA(ZZ$, MKD$(DISPO#), 27, 8)
               Call REPLA(ZZ$, MKD$(EXCES#), 35, 8)
               JJ& = JJ& + 1
               Call GRAREG("LILICRE", ZZ$, JJ&)
            End If
         End If
      End If
   Next IL&
   '
   Call SETULTREG("LILICRE", JJ&)
   Call CIERRARCH("LILICRE")
   '
   OPSALI$ = "IMPRE"
   On Error GoTo 40
   If Option1(0).Value = True Then OPSALI$ = "SCREEN"
   GoTo 41
   '
40 Resume 41
41 On Error GoTo 50
   If Option1(0).Value = True Then OPSALI$ = "SCREEN"
   GoTo 51
   '
50 Resume 51
51 On Error GoTo 0
   If OPSALI$ = "SCREEN" Then
        Call CONECRUN("*COLICRE", "Consulta de Límites de Crédito")
      Else
        Call CONECRUN("*LILICRE", "Listado de Límites de Crédito")
   End If
   Screen.MousePointer = 1
   '
End Sub

Sub GELIGEPRO()
    '
    Dim TICUE%(32767)
    Dim DETICUE$(2)
    Screen.MousePointer = 11
    '
    For U& = 1 To 32767: TICUE%(U&) = 1: Next U&
    DETICUE$(0) = "Cuenta Corriente"
    DETICUE$(1) = "Gastos"
    FIN& = ULTREG&("ACREDOR")
    '
    For U& = 1 To FIN&
        Call TRACE(20, U&, FIN&)
        NCX% = CVI(Left$(REGLEIDO$("ACREDOR", U&), 2))
        If NCX% > 0 Then
            TICUE%(NCX%) = 0
        End If
    Next U&
    '
    FIN& = ULTREG&("PROVED1")
    ULI& = 0
    '
    For U& = 1 To FIN&
        '
        Call TRACE(20, U&, FIN&)
        X$ = REGLEIDO$("PROVED1", U&)
        NP% = CVI(Left$(X$, 2))
        NC% = (-1) * NP%
        '
        Y$ = REGBLAN$("LIGEPRO")
        Call REPLA(Y$, MKI$(NP%), 1, 2)
        CUECOX% = TICUE%(NP%)
        Call REPLA(Y$, DETICUE$(CUECOX%), 3, 30)
        ULI& = ULI& + 1
        Call GRAREG("LIGEPRO", Y$, ULI&)
        '
    Next U&
    '
    Call SETULTREG("LIGEPRO", ULI&)
    Call CIERRARCH("LIGEPRO")
    Screen.MousePointer = 1
    Call FINAL("*LIGEPRO")
    '
End Sub
'

