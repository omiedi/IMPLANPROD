VERSION 5.00
Begin VB.Form ANUFORMACC 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Anulacion de Comprobantes"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   4140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "  Cancela Operación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2140
      Picture         =   "ANUFORMACC.frx":0000
      TabIndex        =   8
      ToolTipText     =   "Cancela la Operación de Anulación"
      Top             =   3780
      Width           =   1800
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00E0E0E0&
      Caption         =   " Confirma Anulación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   210
      Picture         =   "ANUFORMACC.frx":030A
      TabIndex        =   7
      ToolTipText     =   "Anula el Número de Comprobante Mostrado"
      Top             =   3780
      Width           =   1800
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Número de Comprobante"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   210
      TabIndex        =   4
      Top             =   2520
      Width           =   3750
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1680
         TabIndex        =   10
         Top             =   450
         Width           =   1380
      End
      Begin VB.Label Label12 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0001"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   345
         Left            =   840
         TabIndex        =   6
         Top             =   450
         Width           =   690
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1515
         TabIndex        =   5
         Top             =   420
         Width           =   150
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Tipo de Comprobante"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   210
      TabIndex        =   0
      Top             =   315
      Width           =   3750
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Recibos Pre-Numerados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   420
         TabIndex        =   11
         Top             =   1680
         Width           =   2850
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Facturas de Exportación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   420
         TabIndex        =   9
         Top             =   1050
         Width           =   3060
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Facturas, Débitos, Créditos 'B'"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   420
         TabIndex        =   3
         Top             =   735
         Width           =   3060
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Remitos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   420
         TabIndex        =   2
         Top             =   1365
         Width           =   2850
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Facturas, Débitos, Créditos 'A'"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   420
         TabIndex        =   1
         Top             =   420
         Value           =   -1  'True
         Width           =   3060
      End
   End
End
Attribute VB_Name = "ANUFORMACC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub NUFACTURA()
   '
   Static FLINUMERA%
   Static NUMECORR%
   '
   Screen.MousePointer = 11
   If FLINUMERA% = 0 Then
      FLINUMERA% = FECHANUM("01/01/99")
      If ULTREG&("FORMFAC") > 0 Then
         FFS$ = REGLEIDO$("FORMFAC", 2&)
         FLINUMERA% = CVI(Mid$(FFS$, 19, 2))
      End If
   End If
   FLINUME = CDbl(CVDAT(FLINUMERA%))
   '
   NUMECORR% = 1
   VARIDOC% = 1
   If Option1(1).Value = True Then
        VARIDOC% = 2
        VAMOVIX$ = "B"
   End If
   If Option1(2).Value = True Then VARIDOC% = 3
   '
   If Option1(3).Value = True Then
      NROPRO& = 0
      If FORCARGA$ <> "SINREMITO" Then Call NUREMIFA(NROPRO&)
      NUMAXI& = NROPRO& - 1
      GoTo 99
   End If
   '
   If Option1(4).Value = True Then
      NROPRO& = 0
      Call NURECIBO(NROPRO&)
      NUMAXI& = NROPRO& - 1
      GoTo 99
   End If
   '
   If Option1(0) Then
     VAMOVIX$ = "A"
   End If
   
20 Call APERBASDAT("CABAN")
   ASELE$ = "SELECT NumeComp FROM CajaBanc "
   ASELE$ = ASELE$ + "WHERE TipoMovim = 'FVEN' "
   If NUMECORR% <> 1 Then
       ASELE$ = ASELE$ + "AND OpciMovim = '" & OPMOVIX$ & "' "
     Else
       ASELE$ = ASELE$ + "AND (OpciMovim = 'FC' OR OpciMovim = 'ND' OR OpciMovim = 'NC' OR OpciMovim = 'ANUL' ) "
       'ESTA CONDICION LA AGREGUE  PARA QUE TOME EN CTA.
       'LOS Nº DE COMPROBANTES ARRUINADOS.
   End If
   ASELE$ = ASELE$ + "AND VariMovim = '" & VAMOVIX$ & "' "
   If XVALO(Label12) <= 1 Then
       ASELE$ = ASELE$ + "AND (NumeTalo = 0 OR NumeTalo = 1) "
     Else
       ASELE$ = ASELE$ + "AND NumeTalo = " & XVALO(Label12) & " "
   End If
   ASELE$ = ASELE$ + "AND (CDBL(FechEmisi) >= " & FLINUME & ") "
   ASELE = ASELE$ & "ORDER BY NumeComp ASC;"
   '
   Set NUFAXI = CueCorri.OpenRecordset(ASELE$, 4) ' dbOpenDynaset, dbReadOnly)
   '
   NUMAXI& = 0
   On Error Resume Next
   NUFAXI.MoveLast
   If Err < 1 Then
     NUMAXI& = NUFAXI!NUMECOMP
   End If
   '
98 If Visible = True Then
     If NUMAXI& < 1 Then
       NUMAXI& = XVALO(InputBox$("Ingrese Número del Próximo Comprobante del Tipo Elegido", "Incicialización de Talonarios")) - 1
       If NUMAXI& < 0 Then GoTo 99
     End If
   End If
   '
99 NPP$ = TRIM$(Str$(NUMAXI& + 1))
   While Len(NPP$) < 8: NPP$ = "0" + NPP$: Wend
   Label1 = TRIM$(NPP$)
   Label1.Refresh
   '
   Screen.MousePointer = 1
   '
End Sub
'
Sub NURECIBO(NROPRO&)
    '
    Screen.MousePointer = 11
    '
    FENUE$ = CONTROL$("", "FENURECI")
    FLIMI = CVDAT(FECHANUM(FENUE$))
    If FLIMI > Now Then FLIMI = CVDAT(HOY(HO$))
    '
    AUTONURE% = 0
    If CONTROL$("", "NURECIST") = "AUTONUME" Then
      AUTONURE% = 1
    End If
    '
    COBRACT% = 0
    On Error Resume Next
    COBRACT% = XVALO(Label12.Caption)
    On Error GoTo 0
    '
    NROMA& = 0
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT NumeComp, CodiCom_Lar FROM CajaBanc "
    SQLX$ = SQLX$ + "WHERE TipoMovim = 'RCOB' "
    SQLX$ = SQLX$ + "AND CodiCom_Cor LIKE 'RB*' "
    SQLX$ = SQLX$ + "AND VAL(MID$(CodiCom_Lar,6,4)) = " & COBRACT% & " "
    SQLX$ = SQLX$ + "AND Status < 9 "
    SQLX$ = SQLX$ + "AND FechEmisi >= " & FLIMI & " "
    SQLX$ = SQLX$ + "ORDER BY NumeComp DESC;"
    Set OrpaTemp = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    On Error Resume Next
    With OrpaTemp
      .MoveFirst
      If Err < 1 Then
        NROMA& = !NUMECOMP
      End If
      '
      .FindFirst ("NumeComp = " & NROPRO& & " AND VAL(MID$(CodiCom_Lar,6,4)) = " & COBRACT%)
      If .NoMatch = False Then
        NROPRO& = 0
      End If
    End With
    On Error GoTo 0
    '
90  If NROPRO& < 1 Then NROPRO& = 1 + NROMA&
    NPP$ = TRIM$(Str$(NROPRO&))
    While Len(NPP$) < 8: NPP$ = "0" + NPP$: Wend
    Label1 = TRIM$(NPP$)
    Label1.Refresh
    '
    Screen.MousePointer = 1
    '
End Sub
'
Private Sub Command1_Click()
   Call CIERRARCH("*.*")
   Unload ANUFORMACC
End Sub

Private Sub Command13_Click()
   '
   Command13.Enabled = False
   VARIDOC% = 1
   NUCO& = XVALO(Label1)
   FF% = HOY(HOS$)
   '
   Screen.MousePointer = 11
   If Option1(1).Value = True Then VARIDOC% = 2
   If Option1(2).Value = True Then VARIDOC% = 3
   '
   If Option1(3).Value = True Then
      CI% = 0
      CMOV$ = TRIM$(CONTROL$("", "CODREFAC"))
      If CMOV$ = "" Then CMO$ = "RT"
      LOTE$ = ""
      DOPRI$ = CMOV$ + TRIM$(Str$(NUCO&))
      '\\\OT-05-0483-WAR-14/07/05///
        DO1$ = CMOV$ + "-X-"
        DOPRILA1$ = DO1$ + Label12 + "-" + Label1 'RT-X-0001-00000121
      '\\\OT-05-0483-WAR-FIN///
      REFE$ = "Anulacion de Comprobante"
      CANTAA = 1
      Call MOVISTOK(FF%, CI%, LOTE$, CMOV$, DOPRILA1$, DOPRI$, DOSEC$, DOSEC$, REFE$, 0, PRE#, 1, NC%, DEPOSALI%, CANTAA)
      Call CIERRARCH("*.*")
      'NUEVO PARA AVISAR QUE SE ANULO UN COMPROBANTE ARRUINADO.
      Call COMUNI("Se ha Anulado el Comprobante Arruinado\Nº " & Label12 & "-" & Label1 & "\con Reserva de Número")
      Call NUFACTURA
      GoTo 99
   End If
   '
   If Option1(4).Value = True Then    ' anulacion de recibo reservando numero
      REFE$ = "Anulacion de Comprobante"
      CoComCor$ = "RB" & "-" & TRIM$(Str$(XVALO(Label1))) & "-" & Mid$(Label12, 3, 2) 'SIN PREFIJO
      CoComLar$ = "RB" & "-X-" & TRIM$((Label12)) + "-" + TRIM$((Label1))
      CodiEmp% = 0
      CodiSuc% = 0
      NumeClie% = GRUDIVCLI%(0)
      NumeProv% = 0
      FechEmi = CVDAT(FF%)
      Referen$ = "ANULACION"
      FeContab = CVDAT(FF%)
      FeAcredi = CVDAT(FF%)
      CuentaCo% = CUECOCLI%(0)
      TiMovi$ = "RCOB"
      OpMovi$ = "ANUL"
      VaMovi$ = ""
      Descrip$ = REFE$
      NumeTal = XVALO(Label12)
      NumeCom = XVALO(Label1)
      StatInit = 0
      ImpoDeb = 0
      Observa$ = RASOCLI$(0)
      Call GRACABAN
      Call COMUNI("Se ha Anulado el Comprobante Arruinado\Nº " & Label12 & "-" & Label1 & "\con Reserva de Número")
      Call NUFACTURA
      
      GoTo 99
   End If
   '
   'Guardar estos datos en cajabanc
   If Option1(0) Then VARIMOVS$ = "A"
   If Option1(1) Then VARIMOVS$ = "B"
   If Option1(2) Then VARIMOVS$ = "E"
   REFE$ = "Anulacion de Comprobante"
   CoComCor$ = "FC" & "-" & TRIM$(XVALO(Label1)) & "-" & VARIMOVS$ 'SIN PREFIJO
   CoComLar$ = "FC" & "-" & TRIM$((Label12)) + "-" + TRIM$((Label1)) & "-" & VARIMOVS$
   CodiEmp% = 0
   CodiSuc% = 0
   NumeClie% = GRUDIVCLI%(0)
   NumeProv% = 0
   FechEmi = CVDAT(FF%)
   Referen$ = "ANULACION"
   FeContab = CVDAT(FF%)
   FeAcredi = CVDAT(FF%)
   CuentaCo% = CUECOCLI%(0)
   TiMovi$ = "FVEN"
   OpMovi$ = "ANUL"
   VaMovi$ = VARIMOVS$
   Descrip$ = REFE$
   NumeTal = XVALO(Label12)
   NumeCom = NUCO&
   StatInit = 0
   ImpoDeb = 0
   Observa$ = RASOCLI$(0)
   Call GRACABAN
   'NUEVO PARA AVISAR QUE SE ANULO UN COMPROBANTE ARRUINADO.
   Call COMUNI("Se ha Anulado el Comprobante Arruinado\Nº " & Label12 & "-" & Label1 & "\con Reserva de Número")
   Call NUFACTURA
   '
99 Screen.MousePointer = 1
   Command13.Enabled = True
   '
End Sub

Private Sub Form_Load()
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Me.BackColor = &HFCD7B6
      Me.Frame1.BackColor = &HFCD7B6
      Me.Frame9.BackColor = &HFCD7B6
    End If

   Call NUFACTURA
End Sub

Private Sub Label12_Click()
   If Label12 = "0000" Then
       Label12 = "0001"
     ElseIf Label12 = "0001" Then
       Label12 = "0002"
     ElseIf Label12 = "0002" Then
       Label12 = "0003"
     ElseIf Label12 = "0003" Then
       Label12 = "0004"
     Else
       Label12 = "0000"
   End If
   Call NUFACTURA
End Sub

Private Sub Option1_Click(Index As Integer)
   Label12 = "0001"
   If Index = 4 Then
     Label12 = "0000"
   End If
   Call NUFACTURA
End Sub
