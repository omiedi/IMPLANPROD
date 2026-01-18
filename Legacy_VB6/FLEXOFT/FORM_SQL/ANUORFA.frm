VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form ANUORFA 
   BackColor       =   &H00EEEEDD&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Anulación de O-Fabricación"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5940
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   330
      Index           =   0
      Left            =   0
      ScaleHeight     =   270
      ScaleWidth      =   1110
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2160
      TabIndex        =   0
      ToolTipText     =   "Ingreso Manual o por Lectora"
      Top             =   1920
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00800000&
      Height          =   2280
      Left            =   5040
      ScaleHeight     =   2220
      ScaleWidth      =   870
      TabIndex        =   1
      Top             =   0
      Width           =   930
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
         Left            =   105
         Picture         =   "ANUORFA.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   120
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "ANUORFA.frx":014A
         Top             =   1680
         Width           =   1515
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "ANUORFA.frx":206C
      Top             =   -120
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   -360
      OleObjectBlob   =   "ANUORFA.frx":22A0
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo de Conjunto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   960
      TabIndex        =   10
      Top             =   550
      Width           =   1965
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "O-Fabricación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   1440
      TabIndex        =   9
      Top             =   200
      Width           =   1245
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   4815
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2880
      TabIndex        =   7
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2880
      TabIndex        =   6
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Etiqueta de O-F >>>>>"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   2205
   End
End
Attribute VB_Name = "ANUORFA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command2_Click()
  '
  Command2.Enabled = False
  Hide
  Command2.Enabled = True
  '
End Sub


Private Sub Form_Load()
    '
    Call APLICACIONSKINS(Me, Picture1)
    '
End Sub

Private Sub List1_Click()
    '
    Text1.SetFocus
    '
End Sub

Sub ANULAOFABRIC(NORFA$)
     '
     Call ABREORFAB
     Call ABRESERVA

     SQLX$ = "SELECT Statuorf, IdSubOr FROM OrdeFabr "
     SQLX$ = SQLX$ + "WHERE IDSUBOR= '" & NORFA$ & "'"
     Dim ESTADORDEN As ADODB.Recordset
     Set ESTADORDEN = New ADODB.Recordset
     ESTADORDEN.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
     With ESTADORDEN
          On Error Resume Next
          .MoveFirst
          If Err < 1 Then
             Do While Not .EOF
              !StatuOrf = 9
              .Update
              IDSUB$ = SAFETEXT$(!idsubor)
              '
              SQLX$ = "SELECT CantRes, CANTCONS, Saldo_Entre, COD_INTE, IdenLote FROM Reserva "
              SQLX$ = SQLX$ + "WHERE IDSUBOR= '" & IDSUB$ & "' "
              Dim CANTIRES As ADODB.Recordset
              Set CANTIRES = New ADODB.Recordset
1420          On Error Resume Next
              CANTIRES.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
              '
              If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
                On Error GoTo 0
                Call CapturaErrorOpen
                GoTo 1420
              End If
              On Error GoTo 0
              '
              With CANTIRES
                  On Error Resume Next
                  .MoveFirst
                  If Err < 1 Then
                    Do While Not .EOF
                      !CANTRES = XVALO(!CANTCONS)
                      !Saldo_Entre = XVALO(!CANTRES) - XVALO(!CANTCONS)
                      .Update
                      '
                      'ACTUALIZO EL SALDO DEL LOTE
                      CI% = XVALO(!cod_inte)
                      LOTE$ = SAFETEXT$(!idenlote)
                      Call ACSALOTE(CI%, LOTE$, "NOMESS")
                      '
                      .MoveNext
                    Loop
                  End If
                  On Error GoTo 0
              End With
              Set CANTIRES = Nothing
              'ANULACION DE OPERACIONES DE LA O-FABRICACION
              SQLX$ = "DELETE FROM OperFab "
              SQLX$ = SQLX$ + "WHERE IdSubOr = '" & IDSUB$ & "' "
              FLEXCONN.Execute (FILTSQL$(SQLX$))
              .MoveNext
             Loop
          End If
          On Error GoTo 0
     End With
     Set ESTADORDEN = Nothing
     '
     Call ACTUREGISTRO("ORFAETI", "STATUS", "IDSUBOR='" & NORFA$ & "' AND STATUS=0", "9", REGAF&)
     '
     TTXYZ$ = Space$(76)
     HOII% = HOY(HOS$)
     Call REPLA(TTXYZ$, HOS$, 1, 8)
     Call REPLA(TTXYZ$, "Anulacion de O/Fabricación", 11, 56)
     Call REPLA(TTXYZ$, USERTER$, 69, 24)
     Call AGREGA_ANOTAORFAB(NORFA$, TTXYZ$)
     Call AGREGA_ANOTAORFAB(NORFA$, String$(88, "="))
     
     OCOM& = XVALO(VALOBADA("ORDEFABR", "NUME_OCOM", "IDSUBOR='" & NORFA$ & "'"))
     If OCOM& > 0 Then
        OPX% = COMALTER%("La O-Compra " & CStr(OCOM&) & " Esta Asociada a Esta O-F.\¿Desea Anular También?\\No, Continuar\Si, Anular")
        If OPX% = 2 Then
           Call ACTUREGISTRO("OCOMDETA", "STAT_OCOM", "NUME_OCOM=" & CStr(OCOM&), 9, REAF&)
           Call REPLA(TTXYZ$, "Anulacion de O/Compra " & CStr(OCOM&), 11, 56)
           Call AGREGA_ANOTAORFAB(NORFA$, TTXYZ$)
           Call AGREGA_ANOTAORFAB(NORFA$, String$(88, "="))
        End If
     End If
     '
End Sub



Sub Text2_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    MODO% = XVALO(Label2)
    
    'ANULACION DE ORDEN DE FABRICACION
    If Len(Text2) = 9 And MODO% = 0 Then
        'OBTENGO NUMERO DE ORDEN DE FABRICACION Y ORDEN DE ETIQUETA
        XX$ = Text2
        NORFA$ = UCase(Mid$(XX$, 1, 2)) & "-" & Mid$(XX$, 3, 6) & "-" & Mid$(XX$, 9, 1)
        NORFA$ = UCase(NORFA$)
        STAT% = XVALO(VALOBADA("ORDEFABR", "StatuOrf", "IDSUBOR='" & NORFA$ & "'"))
        If STAT% = 3 Then
           Call MENSERR(24, "Orden de Fabricación Cerrada")
           Text2.SelStart = 0
           Text2.SelLength = Len(Text2)
           GoTo 60
        ElseIf STAT% > 3 Then
           Call MENSERR(24, "Orden de Fabricación Anulada")
           Text2.SelStart = 0
           Text2.SelLength = Len(Text2)
           GoTo 60
        End If
        '
        CONDI$ = "IDSUBOR='" & NORFA$ & "'"
        If CantRegistros("ORFAETI", CONDI$) > 0 Then ' VERIFICA QUE SE HAYAN GENERADO ETIQUETAS
            CONDI$ = CONDI$ & " AND STATUS <1"
            If CantRegistros("ORFAETI", CONDI$) < 1 Then ' VERIFICAR QUE HAYA ETIQUETAS/CAJAS PENDIENTES DE INGRESAR
               Call MENSERR(24, "La O-F Leída Se Ha Cerrado Completamente")
               Exit Sub
            End If
        End If
        '
        CIXI% = XVALO(VALOBADA("ORDEFABR", "COD_INTE", "IDSUBOR='" & NORFA$ & "'"))
        Label3.Caption = NORFA$
        Label4.Caption = CODEXT$(CIXI%)
        Label5.Caption = DESCRIT0$(CIXI%)
        '
        OPX% = COMALTER%("¿Esta Seguro de Anular la O-F?\\No, Cancelar\Si, Continuar")
        If OPX% = 2 Then
            Call ANULAOFABRIC(NORFA$)
        End If
        GoTo 50
        '
    End If
    
    'ANULACI0N DE CIERRE DE O-FABRICACION
    If Len(Text2) = 13 And MODO% = 1 Then
        'OBTENGO NUMERO DE ORDEN DE FABRICACION Y ORDEN DE ETIQUETA
        XX$ = Text2
        NORFA$ = UCase(Mid$(XX$, 1, 2)) & "-" & Mid$(XX$, 3, 6) & "-" & Mid$(XX$, 9, 1)
        NORFA$ = UCase(NORFA$)
        STAT% = XVALO(VALOBADA("ORDEFABR", "StatuOrf", "IDSUBOR='" & NORFA$ & "'"))
        NUORDEN% = XVALO(Mid$(XX$, 11, 3))
        '
        CONDI$ = "IDSUBOR='" & NORFA$ & "' AND NUORDEN=" & NUORDEN% & " AND STATUS=3"
        If CantRegistros("ORFAETI", CONDI$) < 1 Then
           Call MENSERR(24, "La Etiqueta Leida No Ha Sido Cerrada")
           GoTo 50
        End If
        '
        FF% = CVINT(VALOBADA("ORFAETI", "FECHACIERRE", CONDI$))
        CI% = XVALO(VALOBADA("ORDEFABR", "COD_INTE", "IDSUBOR='" & NORFA$ & "'"))
        '
        If COMALTER%("¿ Confirma la Anulación de Este Cierre ?\\No, Conservarlo\Si, Anularlo") <> 2 Then GoTo 50
        '
        Screen.MousePointer = 11
        FEREAL = VALOBADA("MOVISTO", "FeReMovi", "DOPRICOR='" & NORFA$ & "' AND FECHMOV='" & FETEXCOR1$(FF%) & "' AND CANTINGRE>0")
        CANFAB = XVALO(VALOBADA("MOVISTO", "CANTINGRE", "DOPRICOR='" & NORFA$ & "' AND FEREMOVI='" & FEREAL & "'"))
        LOTIND = LINDIVEN(CI%)
        CANFAB = CANFAB - LOTIND ' CANTIDAD ACTUALIZADA
        'OBTIENE DEPOSITO DE REGRESO DE LOS COMPONENTES
        DP2% = 0
        ARMAD% = XVALO(VALOBADA("ORDEFABR", "ARMADOR", "IdSubOr = '" & NORFA$ & "'", "NOMESS"))
        If ARMAD% > 0 Then
            TX$ = FILTERGETRECORD$("ARMADORS", 1, MKI$(ARMAD%))
            DP2% = Asc(Mid$(TX$, 71, 1))
        Else
            DP2% = XVALO(Control$("", "DEPODEST"))
        End If
        
        'ACTUALIZACION DEL MOVISTO
        Call ACTUREGISTRO("MOVISTO", "CANTINGRE", "DOPRICOR='" & NORFA$ & "' AND FEREMOVI='" & FEREAL & "'", CANFAB, REGAF&)
        '
        Call CARCOMPO(CODEXT$(CI%), "", CDbl(LOTIND), 0, 0)
        For II& = 1 To ULTREG&("!MATEROF")
            XX$ = REGLEIDO$("!MATEROF", II&)
            CIXI% = CVI(Mid$(XX$, 1, 2))
            CANTI = CVS(Mid$(XX$, 5, 4))
            FF1% = HOY(HO$): JJ% = II&
            DOSELA$ = "IF-" & Mid$(NORFA$, 4)
            Call MOVISTOK(FF1%, CIXI%, "", "CF", NORFA$, NORFA$, DOSELA$, DOSELA$, "Anulacion Cierre de O-F", JJ%, 0, 1, 0, DP2%, CANTI)
        Next II&
        '
        'ACTUALIZO LA CANTIDAD APROBADA DE LA ORDEN DE FABRICACION
        CANTI = XVALO(VALOBADA("ORDEFABR", "CanApro", "IDSUBOR='" & NORFA$ & "'"))
        CANTI = CANTI - LOTIND
        Call ACTUREGISTRO("ORDEFABR", "CanApro", "IDSUBOR='" & NORFA$ & "'", CANTI, REGAF&)
        Call ACTUREGISTRO("ORDEFABR", "StatuOrf", "IDSUBOR='" & NORFA$ & "'", 1, REGAF&)
        '
        'ACTUALIZO TABLA DE ETIQUETAS
        Call ACTUREGISTRO("ORFAETI", "STATUS", "IDSUBOR='" & NORFA$ & "' AND NUORDEN=" & NUORDEN%, 0, REGAF&)
        Call ACTUREGISTRO("ORFAETI", "FECHACIERRE", "IDSUBOR='" & NORFA$ & "' AND NUORDEN=" & NUORDEN%, FETEXCOR1$(FECHANUM("01/01/80")), REGAF&)
        Screen.MousePointer = 1
        '
        Call COMUNI("Movimiento Anulado")
        '
    End If
    
50      Text2 = ""
        GoTo 99
60      Text2.SetFocus
        '
99      Exit Sub

End If

End Sub


