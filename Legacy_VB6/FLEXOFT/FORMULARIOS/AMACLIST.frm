VERSION 5.00
Begin VB.Form AMACLIST 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Maestro de Clientes - Consultas y Listados"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   5385
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Postal"
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
      Left            =   4370
      Picture         =   "AMACLIST.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Listado de Direcciones Postales"
      Top             =   1785
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Mails"
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
      Left            =   4370
      Picture         =   "AMACLIST.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Libreta de Direcciones de Mails"
      Top             =   1050
      Width           =   855
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   4370
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   14
      Top             =   3045
      Width           =   855
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.CommandButton command1 
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
      Left            =   4370
      Picture         =   "AMACLIST.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Cierra - Abandona la Aplicación"
      Top             =   315
      Width           =   855
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFC0&
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
      Height          =   3795
      Left            =   160
      TabIndex        =   0
      Top             =   210
      Width           =   4005
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1260
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   18
         Top             =   2100
         Width           =   2325
         Begin VB.CommandButton Command7 
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
            Picture         =   "AMACLIST.frx":075E
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Listado de Lugar de Entrega"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Listado Lugares de Entrega"
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
            TabIndex        =   19
            Top             =   105
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   945
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   1
         Top             =   1575
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
            Left            =   -20
            Picture         =   "AMACLIST.frx":0A68
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Deudas y Límites de Crédito"
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
            TabIndex        =   3
            Top             =   60
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   630
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   7
         Top             =   1050
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
            Picture         =   "AMACLIST.frx":0D72
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Clientes Inactivos / Activos según fecha de corte"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Listados Cuentas Activas/Inactivas"
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
            TabIndex        =   9
            Top             =   50
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   315
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   10
         Top             =   525
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
            Picture         =   "AMACLIST.frx":107C
            Style           =   1  'Graphical
            TabIndex        =   11
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
            TabIndex        =   12
            Top             =   50
            Width           =   1590
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0FFC0&
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
         TabIndex        =   4
         Top             =   2940
         Width           =   3165
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0FFC0&
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
            TabIndex        =   6
            Top             =   250
            Value           =   -1  'True
            Width           =   1275
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
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
            TabIndex        =   5
            Top             =   250
            Width           =   1380
         End
      End
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   3780
      Picture         =   "AMACLIST.frx":1386
      Top             =   3480
      Width           =   2010
   End
End
Attribute VB_Name = "AMACLIST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Call CIERRARCH("*.*")
   Hide
End Sub

Private Sub Command2_Click()
   '
   Command2.Enabled = False
   CMQ% = COMALTER%("Opciones Disponibles\\Listado Impreso\Exportar")
   If CMQ% < 1 Or CMQ% > 2 Then GoTo 99
   '
   Screen.MousePointer = 11
   If CMQ% = 1 Then
      FIN& = ULTREG&("DEUDOR1")
      For U& = 1 To FIN&
        Call TRACE(20, U&, FIN&)
        X$ = REGLEIDO$("DEUDOR1", U&)
        NC% = CVI(Left$(X$, 2))
        EMA$ = EMAILCLI$(NC%)
        EMB$ = TEL1CLI$(NC%)
        EMC$ = FAXCLI$(NC%)
        If TRIM$(EMA$ + EMB$ + EMC$) <> "" Then
          TTXX$ = String$(2, 0) + Space$(126)
          Call REPLA(TTXX$, MKI$(NC%), 1, 2)
          Call REPLA(TTXX$, EMB$, 33, 24)
          Call REPLA(TTXX$, EMC$, 57, 24)
          Call REPLA(TTXX$, EMA$, 81, 48)
          JJ& = JJ& + 1
          Call GRAREG("LITELMAI", TTXX$, JJ&)
        End If
      Next U&
      '
      Call SETULTREG("LITELMAI", JJ&)
      Call CIERRARCH("LITELMAI")
      Call CIERRARCH("*.*")
      '
      Call FINAL("*LITELMAI")
      GoTo 99
      '
   End If
   '
   OPX% = COMALTER%("Opciones de Exportación:\\Solo E-Mails\Teléfonos y Mails")
   AORI$ = "DEUDOR": SIG% = 1
   '
10 FIN& = ULTREG&(AORI$ + "1")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     X$ = REGLEIDO$(AORI$ + "1", U&)
     NC1% = CVI(Left$(X$, 2))
     If NC1% > 0 Then
       NC% = SIG% * NC1%
       EMA$ = EMAILCLI$(NC%)
       EMB$ = TEL1CLI$(NC%)
       EMC$ = FAXCLI$(NC%)
       EMX$ = EMA$: If OPX% = 2 Then EMX$ = TRIM$(EMA$) + TRIM$(EMB$) + TRIM$(EMC$)
       If TRIM$(EMX$) <> "" Then
         If (InStr(EMA$, "@") > 0 Or OPX% = 2) Then
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
     End If
   Next U&
   '
   Call CIERRARCH("*.*")
   If AORI$ = "DEUDOR" Then
     AORI$ = "PROVED"
     SIG% = (-1)
     GoTo 10
   End If
   '
   Call SETULTREG("!BCONTACT", JJ&)
   Call CIERRARCH("!BCONTACT")
   '
99 Command2.Enabled = True
   Screen.MousePointer = 1
   End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   Call LILIMICRE
   Command3.Enabled = True
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   '
   Screen.MousePointer = 11
   FIN& = ULTREG&("DEUDOR1")
   JJ& = 0
   Call COPYSTRU("BADACLIE", "BADACLIX")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     X$ = REGLEIDO$("DEUDOR1", U&)
     NC% = CVI(Left$(X$, 2))
     If NC% > 0 Then
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
   Next U&
   '
   Call SETULTREG("BADACLIX", JJ&)
   Call CIERRARCH("BADACLIX")
   Call CIERRARCH("*.*")
   '
   Call FINAL("*LIPOSTAL")
   Screen.MousePointer = 1
   '
   Command4.Enabled = True
End Sub

Private Sub Command5_Click()
'Exit Sub
Dim ACTI%(32767)
    Command5.Enabled = False
    '\\\OT-06-0194-WAR-29/07/06///
10  OPX1% = COMALTER%("Seleccione el Listado:\\Clientes Inactivos\Clientes Activos")
    If OPX1% = 2 Then
      '
      HO1% = HOY(HO$)
      Call REPLA(VDEF$, MKI$(HO1%), 1, 2)
      OBJ1$ = OBJPLA$("ENTRAFECHA1", VDEF$)
      If TRIM$(OBJ1$) = "" Then GoTo 10
      '
      FLIM% = CVI(Left$(OBJ1$, 2))
      If FLIM% > 0 Then
         Screen.MousePointer = 11
         FIN& = ULTREG&("FACTUR")
         For U& = 1 To FIN&
           Call TRACE(20, U&, FIN&)
           X$ = REGLEIDO$("FACTUR", U&)
           NC% = CVI(Left$(X$, 2))
           F1% = CVI(Mid$(X$, 7, 2))
           If F1% >= FLIM% Then
             If NC% <= 32767 Then
               If NC% > 0 Then
                 ACTI%(NC%) = F1%
               End If
             End If
           End If
         Next U&
         '
         FIN& = ULTREG&("COBRAN")
         For U& = 1 To FIN&
           Call TRACE(20, U&, FIN&)
           X$ = REGLEIDO$("COBRAN", U&)
           NC% = CVI(Left$(X$, 2))
           F1% = CVI(Mid$(X$, 7, 2))
           If F1% >= FLIM% Then
             If NC% <= 32767 Then
               If NC% > 0 Then
                 If F1% > ACTI%(NC%) Then
                   ACTI%(NC%) = F1%
                 End If
               End If
             End If
           End If
         Next U&
         '
         FIN& = ULTREG&("PRESUPU")
         For U& = 1 To FIN&
           Call TRACE(20, U&, FIN&)
           X$ = REGLEIDO$("PRESUPU", U&)
           NC% = CVI(Mid(X$, 7, 2))
           F1% = CVI(Mid$(X$, 5, 2))
           If F1% >= FLIM% Then
             If NC% <= 32767 Then
               If NC% > 0 Then
                 If F1% > ACTI%(NC%) Then
                  ACTI%(NC%) = F1%
                 End If
               End If
             End If
           End If
         Next U&
         '
         FIN& = ULTREG&("PEDENCA")
         For U& = 1 To FIN&
           Call TRACE(20, U&, FIN&)
           X$ = REGLEIDO$("PEDENCA", U&)
           NC% = CVI(Mid(X$, 55, 2))
           F1% = CVI(Mid$(X$, 53, 2))
           If F1% >= FLIM% Then
             If NC% <= 32767 Then
               If NC% > 0 Then
                 If F1% > ACTI%(NC%) Then
                   ACTI%(NC%) = F1%
                 End If
               End If
             End If
           End If
         Next U&
         '
       RELISA& = 0
       For JX% = 1 To 32767
         Call TRACE(20, IL&, 32767)
         If JX% = 32767 Then Exit For
         NC% = JX%
         If ACTI%(NC%) > 0 Then
           Y$ = REGBLAN$("LISACLI")
           SSCC# = SALDCLI#(NC%)
           Call REPLA(Y$, MKI$(NC%), 1, 2)
           Call REPLA(Y$, MKD$(SSCC#), 3, 8)
           Call REPLA(Y$, MKD$(SSCC#), 11, 8)
           TTXX$ = "Último Movimiento: " + FECHATEX$(ACTI%(NC%))
           Call REPLA(Y$, TTXX$, 19, 30)
           '
           RELISA& = RELISA& + 1
           Call GRAREG("LISACLI", Y$, RELISA&)
         End If
       Next JX%
        '
        Call SETULTREG("LISACLI", RELISA&)
        Call CIERRARCH("*.*")
        Screen.MousePointer = 1
        Call CONECRUN("*DEUACTI", "Listado de Cuentas Activas")
        '
      End If
     '
    ElseIf OPX1% = 1 Then
    '\\\OT-06-0194-WAR-FIN///
      '
      HO1% = HOY(HO$)
      Call REPLA(VDEF$, MKI$(HO1%), 1, 2)
      OBJ1$ = OBJPLA$("ENTRAFECHA1", VDEF$)
      If TRIM$(OBJ1$) = "" Then GoTo 10
      '
      FLIM% = CVI(Left$(OBJ1$, 2))
      If FLIM% > 0 Then
         Screen.MousePointer = 11
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
           If NC% > 0 Then
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
           End If
         Next IL&
         '
         Call SETULTREG("LISACLI", RELISA&)
         Call CIERRARCH("*.*")
         Screen.MousePointer = 1
         Call CONECRUN("*DEUINAC", "Listado de Cuentas Inactivas")
      End If
    End If
99  Screen.MousePointer = 1
    Command5.Enabled = True
End Sub

Private Sub Command7_Click()
  '\\\OT-06-0222-WAR-02/06/06///
  Call FILESORT("LENTRECL", "LENTRECL", 1, 1, "ASC")
  Call FINAL("*LISDAEN")
  '\\\OT-06-0222-WAR-FIN///
End Sub

Private Sub Command8_Click()
    If Option1(0).Value = True Then
        Option2.Value = True
    End If
    Command8.Enabled = False
    FORLICLI.Show 1
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

