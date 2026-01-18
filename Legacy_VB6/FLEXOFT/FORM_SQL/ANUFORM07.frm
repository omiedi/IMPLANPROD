VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form ANUFORM07 
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Anulación de Comprobantes"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   1065
   ClientWidth     =   4005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   4005
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   12
      Top             =   420
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cancela Operación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2040
      Picture         =   "ANUFORM07.frx":0000
      TabIndex        =   8
      ToolTipText     =   "Cancela la Operación de Anulación"
      Top             =   3360
      Width           =   1800
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Confirma Anulación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   120
      Picture         =   "ANUFORM07.frx":030A
      TabIndex        =   7
      ToolTipText     =   "Anula el Número de Comprobante Mostrado"
      Top             =   3360
      Width           =   1800
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2D3C0&
      Caption         =   "NÚMERO DE COMPROBANTE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   935
      Left            =   120
      TabIndex        =   4
      Top             =   2280
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
         Top             =   370
         Width           =   1380
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0001"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   840
         TabIndex        =   6
         Top             =   370
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
         Top             =   340
         Width           =   150
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00E2D3C0&
      Caption         =   "TIPO DE COMPROBANTE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3750
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Recibos Pre-Numerados"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   4
         Left            =   420
         TabIndex        =   11
         Top             =   1700
         Width           =   2850
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Facturas de Exportación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   420
         TabIndex        =   9
         Top             =   1030
         Width           =   3060
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Facturas, Débitos, Créditos 'B'"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   420
         TabIndex        =   3
         Top             =   695
         Width           =   3060
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Remitos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   3
         Left            =   420
         TabIndex        =   2
         Top             =   1365
         Width           =   2850
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Facturas, Débitos, Créditos 'A'"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   420
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   3060
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "ANUFORM07.frx":0614
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   60
      OleObjectBlob   =   "ANUFORM07.frx":0848
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "ANUFORM07"
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
   FLINUME = FETEXCOR1$(FLINUMERA%)
   '
   NUMECORR% = 1
   VARIDOC% = 1
   If Option1(1).Value = True Then
        VARIDOC% = 2
        VAMOVIX$ = "B"
   End If
   If Option1(2).Value = True Then
        VARIDOC% = 3
        VAMOVIX$ = LETREXPO
   End If
   '
   If Option1(3).Value = True Then
      NROPRO& = 0
      Call NUREMIFA(NROPRO&)
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
   'ASELE$ = "SELECT NumeComp FROM CajaBanc "
   ASELE$ = "SELECT max(NumeComp) as maximo FROM CajaBanc "
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
   
   ASELE$ = ASELE$ + "AND FechEmisi >= '" & FLINUME & "' "
   NUMAXI& = 0
   '
   Dim NUFAXI As ADODB.Recordset
   Set NUFAXI = FLEXCONN.Execute(FILTSQL$(ASELE$))
   If Not IsNull(NUFAXI!maximo) Then
        NUMAXI& = NUFAXI!maximo
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
    ' ver ernesto temas fechas
    '
    Screen.MousePointer = 11
    '
    FENUE$ = CONTROL$("", "FENURECI")
    FLIMI = FENUE$
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
    
'    sqlx$ = "SELECT NumeComp, CodiCom_Lar FROM CajaBanc "
'    sqlx$ = sqlx$ + "WHERE TipoMovim = 'RCOB' "
'    sqlx$ = sqlx$ + "AND CodiCom_Cor LIKE 'RB*' "
'    sqlx$ = sqlx$ + "AND VAL(MID$(CodiCom_Lar,6,4)) = " & COBRACT% & " "
'    sqlx$ = sqlx$ + "AND Status < 9 "
'    sqlx$ = sqlx$ + "AND FechEmisi >= " & FLIMI & " "
'    sqlx$ = sqlx$ + "ORDER BY NumeComp DESC;"
'    Set OrpaTemp = CueCorri.OpenRecordset(sqlx$, dbOpenSnapshot)
'    On Error Resume Next
'    With OrpaTemp
'      .MoveFirst
'      If Err < 1 Then
'        NROMA& = !NumeComp
'      End If
'      '
'      .FindFirst ("NumeComp = " & NROPRO& & " AND VAL(MID$(CodiCom_Lar,6,4)) = " & COBRACT%)
'      If .NoMatch = False Then
'        NROPRO& = 0
'      End If
'    End With
'    On Error GoTo 0
    '

   SQLX$ = "SELECT max(NumeComp) as maximo FROM CajaBanc "
   SQLX$ = SQLX$ + "WHERE TipoMovim = 'RCOB' "
   SQLX$ = SQLX$ + "AND CodiCom_Cor LIKE 'RB%' "
   SQLX$ = SQLX$ + "AND FechEmisi >= '" & FLIMI & "'"
   SQLX$ = SQLX$ + "AND cast(substring(CodiCom_Lar,6,4)as int) = " & COBRACT%
   SQLX$ = SQLX$ + "AND Status < 9 "
   
   Dim MIRS As ADODB.Recordset
   Set MIRS = FLEXCONN.Execute(FILTSQL$(SQLX$))
   If Not IsNull(MIRS!maximo) Then
        NROMA& = MIRS!maximo
   End If
   MIRS.Close
   Set MIRS = Nothing

'90
    
    If NROPRO& < 1 Then NROPRO& = 1 + NROMA&
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
   Unload ANUFORM07
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
'      CodiEmp% = 0
'      CodiSuc% = 0
      NumeClie = Int(GRUDIVCLI(0))
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
      Observa$ = rasocli$(0)
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
'   CodiEmp% = 0
'   CodiSuc% = 0
   NumeClie = Int(GRUDIVCLI(0))
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
   Observa$ = rasocli$(0)
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
   UTILIZA_SKIN% = 1
   Call APLICACIONSKINS(Me)
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
   If Index = 2 Then Label12 = PREFEXPO
   If Index = 4 Then
     Label12 = "0000"
   End If
   Call NUFACTURA
End Sub
