VERSION 5.00
Begin VB.Form AMACLIS04 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Maestro de Clientes  - Consultas y Listados"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   5235
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      BackColor       =   &H80000016&
      Height          =   645
      Left            =   735
      ScaleHeight     =   585
      ScaleWidth      =   2265
      TabIndex        =   14
      Top             =   4935
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
         Picture         =   "AMACLIS04.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
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
         TabIndex        =   16
         Top             =   50
         Width           =   1590
      End
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
      Height          =   4050
      Left            =   160
      TabIndex        =   0
      Top             =   210
      Width           =   4005
      Begin VB.PictureBox Picture7 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1260
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   23
         Top             =   2445
         Width           =   2325
         Begin VB.CommandButton Command6 
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
            Picture         =   "AMACLIS04.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Clientes Inactivos / Activos según fecha de corte"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label3 
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
            TabIndex        =   25
            Top             =   105
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1050
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   1
         Top             =   1935
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
            Picture         =   "AMACLIS04.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Listado de Deudas y Límite de Crédito"
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
            Top             =   50
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   840
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   19
         Top             =   1410
         Width           =   2325
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
            Height          =   600
            Left            =   0
            Picture         =   "AMACLIS04.frx":091E
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Libreta de Direcciones de Mails"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Lista de Direccio- nes de Mail"
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
            TabIndex        =   20
            Top             =   50
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   630
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   17
         Top             =   885
         Width           =   2325
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
            Picture         =   "AMACLIS04.frx":0C28
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Listado de Direcciones Postales"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Listado de Domi- cilios Postales"
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
            TabIndex        =   18
            Top             =   50
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   420
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   7
         Top             =   360
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
            Picture         =   "AMACLIS04.frx":0F32
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
         Top             =   3240
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
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00006060&
      Height          =   4530
      Left            =   4305
      ScaleHeight     =   4470
      ScaleWidth      =   1005
      TabIndex        =   10
      Top             =   0
      Width           =   1065
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   12
         Top             =   3540
         Width           =   645
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -105
            TabIndex        =   13
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
         Left            =   105
         Picture         =   "AMACLIS04.frx":123C
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Cierra - Abandona la Aplicación"
         Top             =   210
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -210
         Picture         =   "AMACLIS04.frx":1386
         Stretch         =   -1  'True
         Top             =   3825
         Width           =   1410
      End
   End
End
Attribute VB_Name = "AMACLIS04"
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
   Call APERBASDAT("CLIEN")
   SQLX$ = " Select * from deudor12 "
   SQLX$ = SQLX$ + " ORDER BY Nume_Cli "
   Set CLITEMP = Badaclie.OpenRecordset(SQLX$, dbOpenSnapshot)
     With CLITEMP
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
   If CMQ% = 1 Then   'IMPRESION
    With CLITEMP
      .MoveFirst
      Call REGBLAN("LITELMAI")
      JJ& = 0
      Do While .EOF = False
        NC% = !Nume_Cli
        EMA$ = EMAILCLI$(NC%)
        EMB$ = TEL1CLI$(NC%)
        EMC$ = FAXCLI$(NC%)
        If TRIM$(EMA$ + EMB$ + EMC$) <> "" Then
          Call REPLA(TTXX$, MKI$(NC%), 1, 2)
          Call REPLA(TTXX$, EMB$, 33, 24)
          Call REPLA(TTXX$, EMC$, 57, 24)
          Call REPLA(TTXX$, EMA$, 81, 48)
          JJ& = JJ& + 1
          Call GRAREG("LITELMAI", TTXX$, JJ&)
        End If
        .MoveNext
        Loop
    End With
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
   With CLITEMP             'EXPORTACIÓN
     .MoveFirst
     JJ& = 0
     Call REGBLAN("BCONTACT")
     Do While (.EOF = False)
        NC% = !Nume_Cli
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
   MSGXX$ = "Se ha Exportado con Éxito en "
   MSGXX$ = MSGXX$ + LUCOM$ + "BCONTACT.WKF"
   MsgBox (MSGXX$)
   Call CIERRARCH("*.*")
   '
99 Command2.Enabled = True
   Screen.MousePointer = 1
90  End Sub

Sub Command3_Click()
   Command3.Enabled = False
   Call LILIMICRE
   Command3.Enabled = True
End Sub

 Sub Command4_Click()
Command4.Enabled = False
   '
Screen.MousePointer = 11
   '
Call APERBASDAT("CLIEN")
'    FIN& = ULTREG&("DEUDOR1")
'    JJ& = 0
'    Call COPYSTRU("BADACLIE", "BADACLIX")

'Consulta sql
 SQLX$ = " SELECT * FROM DEUDOR12 "
'Filtro unicamente las cuentas corrientes
 'SQLX$ = SQLX$ + " WHERE Tcpr_Cli <> 1 "
 SQLX$ = SQLX$ + " WHERE Stat_Cli > 0 " ' Que no esten  de baja
 SQLX$ = SQLX$ + " ORDER BY Raso_Cli ASC "
 Set CLITEMP = Badaclie.OpenRecordset(SQLX$, dbOpenSnapshot)
     With CLITEMP
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
With CLITEMP
    JJ& = 0
    Call REGBLAN("BADACLIX")
 Do While (.EOF = False)
'    For U& = 1 To FIN&
'      Call TRACE(20, U&, FIN&)
'      X$ = REGLEIDO$("DEUDOR1", U&)
'      NC% = CVI(Left$(X$, 2))
       NC% = !Nume_Cli
      ' NC% = (-1) * NP%
     'If NC% < 0 Then
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
     'End If
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

Private Sub Command6_Click()
    Command5.Enabled = False
    Dim ACTI%(32767)
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
         '
         Call ABRECABAN
         SQLX$ = "SELECT * FROM cajabanc where tipomovim ='FVEN'"
         Dim FACT As Recordset
20       Set FACT = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
         '
         With FACT
            If .BOF = True And .EOF = True Then GoTo 30
            .MoveLast
            FIN& = .RecordCount
            .MoveFirst
            For U& = 1 To FIN&
              Call TRACE(20, U&, FIN&)
              NC% = XVALO(!NuClien)   'CVI(Left$(X$, 2))
              F1% = CVINT(!FechEmisi) 'CVI(Mid$(X$, 7, 2))
              If F1% >= FLIM% Then
                If NC% <= 32767 Then
                  If NC% > 0 Then
                    If F1% > ACTI%(NC%) Then ACTI%(NC%) = F1%
                  End If
                End If
              End If
              .MoveNext
            Next U&
            .Close
         End With
         '
30       Dim COBRAN As DAO.Recordset
         SQLX$ = "SELECT * FROM CAJABANC WHERE TIPOMOVIM='RCOB'"
         Set COBRAN = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
         '
         With COBRAN
            If .BOF = True And .EOF = True Then GoTo 40
            .MoveLast
            FIN& = .RecordCount
            .MoveFirst
            For U& = 1 To FIN&
              Call TRACE(20, U&, FIN&)
              
              NC% = XVALO(!NuClien)
              F1% = CVINT(!FechEmisi)
              If F1% >= FLIM% Then
                If NC% <= 32767 Then
                  If NC% > 0 Then
                    If F1% > ACTI%(NC%) Then ACTI%(NC%) = F1%
                  End If
                End If
              End If
              .MoveNext
            Next U&
            .Close
         End With
         '
40       Call ABREPRESVEN
         '
         With PresuVen
            If .BOF = True And .EOF = True Then GoTo 50
            .MoveLast
            FIN& = .RecordCount
            .MoveFirst
            For U& = 1 To FIN&
              Call TRACE(20, U&, FIN&)
              'X$ = REGLEIDO$("FACTUR", U&)
              NC% = XVALO(!NroClie) 'CVI(Left$(X$, 2))
              F1% = CVINT(!FECHEMIS)
              If F1% >= FLIM% Then
                If NC% <= 32767 Then
                  If NC% > 0 Then
                    If F1% > ACTI%(NC%) Then ACTI%(NC%) = F1%
                  End If
                End If
              End If
              .MoveNext
            Next U&
            .Close
         End With
         '
50       Call ABREPEDCLI
         With PedEnca
            If .BOF = True And .EOF = True Then GoTo 60
            .MoveLast
            FIN& = .RecordCount
            .MoveFirst
            For U& = 1 To FIN&
              Call TRACE(20, U&, FIN&)
              NC% = XVALO(!NroClie)
              F1% = CVINT(!FECHEMIS)
              If F1% >= FLIM% Then
                If NC% <= 32767 Then
                  If NC% > 0 Then
                    If F1% > ACTI%(NC%) Then ACTI%(NC%) = F1%
                  End If
                End If
              End If
              .MoveNext
            Next U&
            .Close
         End With
         '
60       RELISA& = 0
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
             Call REPLA(Y$, Chr$(VENDCLI%(NC%)), 49, 1)
             '
             RELISA& = RELISA& + 1
             Call GRAREG("LISACLI", Y$, RELISA&)
           End If
         Next JX%
         '
         Call SETULTREG("LISACLI", RELISA&)
         Call FILESORT("LISACLI", "", 1, 1, "ASC")
         Call CIERRARCH("*.*")
         Screen.MousePointer = 1
         Call CONECRUN("*DEUACTI", "Listado de Cuentas Activas")
         '
      End If
      '
    ElseIf OPX1% = 1 Then
      '
      HO1% = HOY(HO$)
      Call REPLA(VDEF$, MKI$(HO1%), 1, 2)
      OBJ1$ = OBJPLA$("ENTRAFECHA1", VDEF$)
      If TRIM$(OBJ1$) = "" Then GoTo 10
      '
      FLIM% = CVI(Left$(OBJ1$, 2))
      If FLIM% > 0 Then
         Screen.MousePointer = 11
         Call ABRECABAN
         SQLX$ = "SELECT * FROM cajabanc where tipomovim ='FVEN'"
         'Dim FACT As DAO.Recordset
80       Set FACT = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
         '
         With FACT
            If .BOF = True And .EOF = True Then GoTo 80
            .MoveLast
            FIN& = .RecordCount
            .MoveFirst
            For U& = 1 To FIN&
              Call TRACE(20, U&, FIN&)
              X$ = REGLEIDO$("FACTUR", U&)
              NC% = XVALO(!NuClien)
              If NC% > 0 Then
                If NC% <= 32767 Then
                  F1% = CVINT(!FechEmisi)
                  If F1% > ACTI%(NC%) Then ' If CVI(Mid$(X$, 7, 2)) > ACTI%(NC%) Then
                    If F1% > ACTI%(NC%) Then ACTI%(NC%) = F1% 'CVI(Mid$(X$, 7, 2))
                  End If
                End If
              End If
            .MoveNext
            Next U&
            .Close
         End With
         '
90       RELISA& = 0
         Call ABRECLIEN
         'Se utiliza otro recordset ya que la funcion ZONACLI utiliza el recordset
         'deudor12
         Set Deudor12_aux = Badaclie.OpenRecordset("DEUDOR12", dbOpenSnapshot)
         With Deudor12_aux
            If .BOF = True And .EOF = True Then GoTo 30
            .MoveLast
            CX& = .RecordCount
            .MoveFirst
            For IL& = 1 To CX&
              Call TRACE(20, IL&, CX&)
              NC% = XVALO(!Nume_Cli) 'CVI(Left$(REGLEIDO$("DEUDOR1", IL&), 2))
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
                   Call REPLA(Y$, Chr$(VENDCLI%(NC%)), 49, 1)
                   '
                   RELISA& = RELISA& + 1
                   Call GRAREG("LISACLI", Y$, RELISA&)
                End If
              End If
              .MoveNext
            Next IL&
            .Close
         End With
         '
         Call SETULTREG("LISACLI", RELISA&)
         Call FILESORT("LISACLI", "", 1, 1, "ASC")
         Call CIERRARCH("*.*")
         Screen.MousePointer = 1
         '
         If Option2.Value = True Then
            Call CONECRUN("*LDEUINA", "Listado de Cuentas Inactivas")
          Else
            Call CONECRUN("*DEUINAC", "Consulta de Cuentas Inactivas")
         End If
         '
      End If
    End If
99  Screen.MousePointer = 1
    Command5.Enabled = True

End Sub

Sub Command8_Click()
    If Option1(0).Value = True Then
        Option2.Value = True
    End If
    Command8.Enabled = False
    FORLICLI04.Show 1
    Command8.Enabled = True
End Sub

Sub LILIMICRE()
    '
' falta considerar deuda en cheques en cartera
    '
    Screen.MousePointer = 11
    Dim DEUCUECO#(32767), CHEPEAC#(32767)
    '
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM SADEUPEN "
    'SQLX$ = SQLX$ + "WHERE NuClien = " & NC% & " "
    'SQLX$ = SQLX$ + "AND (TipoMovim = 'FCOM' "
    'SQLX$ = SQLX$ + "OR TipoMovim = 'OPAG') "
    'SQLX$ = SQLX$ + "ORDER BY FeVencim ASC;"
    Set FacPenTemp = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With FacPenTemp
      On Error Resume Next
      .MoveLast
      If Err < 1 Then
        On Error GoTo 0
        .MoveFirst
        Do While Not .EOF
          NCI% = !NuClien
          If NCI% >= 0 Then
            SALDEU# = !SaldDebe - !SaldAcre
            DEUCUECO#(NCI%) = DEUCUECO#(NCI%) + SALDEU#
          End If
        .MoveNext
        Loop
      End If
    End With
    '
'FALTA INCLUIR CHEQUES PENDIENTES DE ACREDITACION
    HOYDIA = CDbl(Int(Now))
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "where TipoMovim = 'CHRECI' "
    SQLX$ = SQLX$ + "AND Status < 9 "
    SQLX$ = SQLX$ + "AND CDBL(FechAcred) >= " & HOYDIA & " "
    '
    Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With CABATEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        On Error GoTo 0
        Do While Not .EOF
          If TRIM$(!OpciMovim) <> "CH" Then GoTo 19 ' no es un cheque
          '
          ICHEQ# = !ValAbsComp
          NCI% = !NuClien
          '
          If NCI% > 0 Then
            If NCI% <= 32767 Then
              CHEPEAC#(NCI%) = CHEPEAC#(NCI%) + ICHEQ#
            End If
          End If
          '
19      .MoveNext
        Loop
        '
      End If
    End With
    '
    FIN& = 32767
    JJ& = 0
    For IL& = 1 To FIN&
      Call TRACE(20, IL&, FIN&)
      'XX$ = REGLEIDO$("DEUDOR1", IL&)
      NCI% = IL&    '   CVI(Left$(XX$, 2))
      'If NCI% > 0 Then
      '   If NCI% <= 16384 Then
      If Abs(DEUCUECO#(NCI%)) + Abs(CHEPEAC#(NCI%)) >= 0.005 Then
        '       YY$ = REGLEIDO$("DEUDOR2", IL&)
        CRELIM# = LIMICRE(NCI%) 'CVS(Mid$(YY$, 121, 4))
        DISPO# = CRELIM# - DEUCUECO#(NCI%) - CHEPEAC#(NCI%)
        EXCES# = 0
        If DISPO# < 0 Then
          EXCES# = Abs(DISPO#)
          DISPO# = 0
        End If
        ZZ$ = REGBLAN$("LILICRE")
        Call REPLA(ZZ$, MKI$(NCI%), 1, 2)
        Call REPLA(ZZ$, MKD$(CRELIM#), 3, 8)
        If CRELIM# >= 0 Then
          Call REPLA(ZZ$, MKD$(DEUCUECO#(NCI%)), 11, 8)
          Call REPLA(ZZ$, MKD$(CHEPEAC#(NCI%)), 19, 8)
          Call REPLA(ZZ$, MKD$(DISPO#), 27, 8)
          Call REPLA(ZZ$, MKD$(EXCES#), 35, 8)
        End If
        JJ& = JJ& + 1
        Call GRAREG("LILICRE", ZZ$, JJ&)
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
40  Resume 41
41  On Error GoTo 50
    If Option1(0).Value = True Then OPSALI$ = "SCREEN"
    GoTo 51
    '
50  Resume 51
51  On Error GoTo 0
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



