VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form SELNUMSE 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selección de Números de Serie "
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   7755
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   7755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   22
      Top             =   720
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.ListBox LISNUM 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2865
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   0
      ToolTipText     =   "Tilde para seleccionar"
      Top             =   960
      Width           =   6465
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00E0E0E0&
      Height          =   435
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   6405
      TabIndex        =   7
      Top             =   495
      Width           =   6465
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "N.Serie          Descripción"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   225
         Left            =   220
         TabIndex        =   8
         Top             =   105
         Width           =   10515
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00400000&
      Height          =   5775
      Left            =   6840
      ScaleHeight     =   5715
      ScaleWidth      =   1455
      TabIndex        =   3
      Top             =   0
      Width           =   1515
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
         Left            =   120
         Picture         =   "SELNUMSE.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   120
         Width           =   650
      End
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
         Height          =   650
         Left            =   120
         Picture         =   "SELNUMSE.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Refresca - Invierte Selección"
         Top             =   2850
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   4
         Top             =   4425
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   5
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton BOTREC 
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
         Left            =   120
         Picture         =   "SELNUMSE.frx":0294
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Aprueba la Selección"
         Top             =   3555
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "SELNUMSE.frx":059E
         Top             =   4770
         Width           =   1515
      End
   End
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   10200
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   5400
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Depósitos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   13
      Top             =   4080
      Width           =   6475
      Begin VB.CommandButton Command6 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2940
         TabIndex        =   15
         Top             =   240
         Width           =   175
      End
      Begin VB.CommandButton Command3 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6060
         TabIndex        =   14
         Top             =   240
         Width           =   175
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dep.Consumo: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4440
         TabIndex        =   21
         Top             =   300
         Width           =   1320
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dep.Entrada: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   840
         TabIndex        =   20
         Top             =   300
         Width           =   1800
      End
      Begin VB.Label Label8 
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
         Height          =   285
         Left            =   240
         TabIndex        =   19
         Top             =   660
         Width           =   2895
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   2640
         TabIndex        =   18
         Top             =   240
         Width           =   315
      End
      Begin VB.Label Label2 
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
         Height          =   285
         Left            =   3360
         TabIndex        =   17
         Top             =   660
         Width           =   2895
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   5760
         TabIndex        =   16
         Top             =   240
         Width           =   315
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "SELNUMSE.frx":24C0
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "SELNUMSE.frx":26F4
      TabIndex        =   23
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nº.O.Fabricación:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3480
      TabIndex        =   12
      Top             =   120
      Width           =   1605
   End
   Begin VB.Label Label34 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   5190
      TabIndex        =   11
      Top             =   120
      Width           =   1395
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   5500
      Y1              =   3255
      Y2              =   3255
   End
   Begin VB.Line Line2 
      X1              =   5505
      X2              =   5505
      Y1              =   495
      Y2              =   3245
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   330
      Left            =   7440
      TabIndex        =   9
      Top             =   1140
      Width           =   1695
   End
End
Attribute VB_Name = "SELNUMSE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'


Private Sub BOTREC_Click()
     BOTREC.Enabled = False
     '
     ' VERIFICO QUE HAYA AL MENOS UNO SELECCIONADO
     CANTIEQUI# = 0
     For II2% = 1 To SELNUMSE.LISNUM.ListCount
        If SELNUMSE.LISNUM.Selected(II2% - 1) = True Then CANTIEQUI# = CANTIEQUI# + 1
     Next II2%
     If CANTIEQUI# < 0.5 Then
        Call MENSERR(24, " Debe Seleccionar los Números de Serie ")
        GoTo 99
    End If
    '
    'VALIDA QUE NO CAMBIE EL DEPOSITO EN EL CASO QUE SEA UN ARMADOR EXTERNO.
    NORFA$ = TRIM$(Label34)
    If ValidaArmador%(NORFA$) > 0 Then
       DEPARMA% = XVALO(Control("", "DEPOARMA"))
       If XVALO(Label3) <> DEPARMA% Then
        Call MENSERR(24, "No Válido Cambiar de Depósito\Orden de Fabricación Asignado a Armador Externo\Depósito de Consumo: " & DEPARMA%)
        GoTo 99
       End If
    End If
    '
    Screen.MousePointer = 11
    CI1% = VALOBADA("ORDEFABR", "Cod_Inte", "IdSubOr = '" & NORFA$ & "'", "NOMESS")
    VACONTROL% = 1
    CODICONJ$ = CODEXT$(CI1%)
50  Call CARCOMPO(CODICONJ$, NORFA$, CANTIEQUI#, VACONTROL%, 3)
    If VACONTROL% < 0 Then GoTo 99                 ' SI CANCELA SE VA
    '
    If Control("VANEDECE", "DEPOCEN") = "SI" Then ' SI CONFIGURADO QUE
       For i& = 1 To ULTREG("!MATEROF")          ' QUE
          XXX$ = REGLEIDO$("!MATEROF", i&)
          CI% = CVI(Mid$(XXX$, 1, 2))
          UNIME$ = Mid$(XXX$, 3, 2)
          UsoUnit# = CVD(Mid$(XXX$, 9, 8))
          NECESI# = UsoUnit# * CANTIEQUI#
          '
          ' stock transferido a deposito taller
          CONDI$ = "Cod_Inte = " & CI% & " AND DOSECCOR = '" & TRIM$(NORFA$) & "' And DepoMovi = " & XVALO(Label3)
          STOXXX1# = XVALO(VALOBADA("MOVISTO", "SUM(CantIngre)", CONDI$, "NOMESS"))
          '
          ' stock consumido de deposito taller
          CONDI$ = "CodiMovi = 'CF' AND Cod_Inte = " & CI% & " AND DOPRICOR = '" & TRIM$(NORFA$) & "' AND DepoMovi = " & XVALO(Label3)
          STOXXX2# = XVALO(VALOBADA("MOVISTO", "SUM(CantSalid)", CONDI$, "NOMESS"))
          '
          ' saldo en taller
          SOTOCKK# = STOXXX1# - STOXXX2#
          If SOTOCKK# < NECESI# - 0.005 Then
             Call MENSERR(24, "Imposible Cerrar Esta O.F.: " & NORFA$ & "\Código: " & CODEXT$(CI%) & " con Faltantes\En Depósito: " & Label3 & "")
             GoTo 99
          End If
          '
       Next i&
    End If
    
    CAN = CANTIEQUI#
    DEPX% = XVALO(SELNUMSE.Label1)   ' DEPOSITO DE DESTINO
    DP2% = XVALO(SELNUMSE.Label3)    ' DEPOSITO DE CONSUMO
    HOII% = HOY(HOS$)
    FECHA% = HOII%
    CI2% = CI1%
    LOTE$ = NORFA$
    CMOV$ = "OF"
    DOPRI$ = NORFA$
    DOSEC$ = "IF-" + TRIM$(Mid$(NORFA$, 4))
    REFE$ = "Fabricación Propia"
    VUNI# = 0
    MONE$ = "$"
    Call MOVISTOK(FECHA%, CI2%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NUORIT%, VUNI#, MONEII%, NC%, DEPX%, CAN)
    '
    'ACTUALIZA LOS DEPOSITOS  EN TANUMSE
    For II2% = 1 To SELNUMSE.LISNUM.ListCount
      If SELNUMSE.LISNUM.Selected(II2% - 1) = True Then
        NROSERIE$ = TRIM$(Mid$(SELNUMSE.LISNUM.List(II2% - 1), 1, 16))
        Call ACTUREGISTRO("TANUMSE", "DEPONUMSE", "NUMEROSERIE='" & NROSERIE$ & "'", XVALO(Label1), REG&)
      End If
    Next II2%
    '
    CANTI = FORMATNUM(CANTIEQUI#, "F12.1") ' GRABA EN EL ANOTADOR
    TTXYZ$ = Space$(76)
    HOII% = HOY(HOS$)
    Call REPLA(TTXYZ$, HOS$, 1, 8)
    Call REPLA(TTXYZ$, "Cierre de O/Fabricación", 11, 56)
    Call REPLA(TTXYZ$, USERTER$, 69, 24)
    Call AGREGA_ANOTAORFAB(NORFA$, TTXYZ$)
    '
    TTXYZ$ = Space$(76)
    Call REPLA(TTXYZ$, "     - Cantidad Aprob.=  " & CANTI, 11, 56)
    Call AGREGA_ANOTAORFAB(NORFA$, TTXYZ$)
    Call AGREGA_ANOTAORFAB(NORFA$, String$(88, "="))
    '
    ' ANOTA SALDO ACUMULADO DE LA O/F
    CONDI$ = "CODIMOVI = 'OF' AND DOPRILAR = '" & NORFA$ & "'"
    TOFABRI = XVALO(VALOBADA("MOVISTO", "SUM(CANTINGRE)", CONDI$, "NOMESS"))
    TOFABRI = TOFABRI - XVALO(VALOBADA("MOVISTO", "SUM(CANTSALID)", CONDI$, "NOMESS"))
    If TOFABRI < 0 Then TOFABRI = 0
    Call ACTUREGISTRO("ORDEFABR", "CanApro", "IdSubOr = '" & NORFA$ & "'", TOFABRI, REGAFE&)
    If REGAFE < 1 Then
       Call MENSERR(24, "Imposible Ajustar Saldo Pendiente\de la Orden de Fabricación '" + NORFA$ + "'.")
    End If
    '
    ' declara consumo de componentes
    FIMATER& = ULTREG&("!MATEROF")
    For U& = 1 To FIMATER&
      Call TRACE(20, U&, FIMATER&)
      ZZ$ = REGLEIDO$("!MATEROF", U&)
      CI2% = CVI(Left$(ZZ$, 2))
      CAN = (-1) * CVS(Mid$(ZZ$, 5, 4))
      If TRIM$(UCase$(Unimed$(CI2%))) = "GK" Then CAN = CAN / 1000
      '
      USOUNI# = CVD(Mid$(ZZ$, 9, 8))
      RERES& = CVS(Mid$(ZZ$, 17, 4))
      If CI2% > 0 Then
        If CI2% <= NUMAS% Then
          If Abs(CAN) >= 0.05 Then
            LOTE$ = ""      ' MÁS ABAJO SE REE DE RESERVAS
            CMOV$ = "CF"
            '
            DOPRI$ = NORFA$
            DOSEC$ = "PF-" + TRIM$(Mid$(NORFA$, 4))
            REFE$ = "CONSUMO " + CODICONJ$
            If CAN > 0 Then REFE$ = "RECUPERO " + CODICONJ$
            '
            DEPX% = DP2%: If DP2% = 99 Then DEPX% = DEPOPRE%(CI2%)
144         NUORIT% = NUORIT% + 1
            '
            ' determina el lote reservado
            CONDIX$ = " IdSubOr = '" & NORFA$ & "' AND Cod_Inte = " & CI2% & ""
            LOTE$ = VALOBADA("Reserva", "IdenLote", CONDIX$, "NOMESS")
            '
            Call MOVISTOK(FECHA%, CI2%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NUORIT%, VUNI#, MONEII%, NC%, DEPX%, CAN)
            CIMOVI%(CI2%) = 1
            '
            ' cancela las reservas del componente
            SALCON# = (-1) * CAN
            SQLX$ = "SELECT * FROM Reserva "
            SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NORFA$ & "' "
            SQLX$ = SQLX$ + " AND Cod_Inte = " & CI2% & " "
            Dim RESETEMP As ADODB.Recordset
            Set RESETEMP = New ADODB.Recordset
            RESETEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
            On Error Resume Next
            With RESETEMP
145           .MoveFirst
              If Err < 1 Then
                On Error GoTo 0
                Do While Not .EOF
                   CARESE# = !CANTRES
                   CACOANT# = !CANTCONS
                   If CARESE# - CACOANT# >= SALCON# Then
                        !CANTCONS = CACOANT# + SALCON#
                        SALCON# = 0
                      Else
                        SALCON# = SALCON# - (CARESE# - CACOANT#)
                        !CANTCONS = !CANTRES
                   End If
                   .Update
                   .MoveNext
                Loop
              End If
              On Error GoTo 0
            End With
            Set RESETEMP = Nothing
            '
          End If
        End If
      End If
    Next U&
    '
    CANPROYE& = XVALO(VALOBADA("ORDEFABR", "CanProy", "IdSubOr = '" & NORFA$ & "'"))
    CANREALE& = XVALO(VALOBADA("ORDEFABR", "CanApro", "IdSubOr = '" & NORFA$ & "'"))
    '
    If CANREALE& >= CANPROYE& Then
        CONDI$ = "IdSubOr = '" & NORFA$ & "'"
        HOII% = HOY(HOS$)
        FECHAHOY = FETEXCOR1$(HOII%)
        Call ACTUREGISTRO("ORDEFABR", "StatuOrf", CONDI$, 3, REGAF&, "NOMESS")
        Call ACTUREGISTRO("ORDEFABR", "FechCierr", CONDI$, FECHAHOY, REGAF&, "NOMESS")
        ' Cancela Operaciones Programadas
        Call BORRAREGISTROS("OperFab", CONDI$, RAFE&, "NOMESS")
        ' Cancela Reservas de Componentes
        Call BORRAREGISTROS("Reserva", CONDI$, RAFE&, "NOMESS")
        ' Cancela Orden en Indice de Ordenes Pendientes
        NORFAX$ = AJUSTI$(NORFA$, 12)
        Call FILTERDELETE("ECORDEP", 1, NORFAX$)
        Call FRESHECHO("ECORDEP")
    End If
    Call CIERRARCH("*.*")
    Call LIBARCH("ECORDEP")
    Call CIERRARCH("*.*")
    '
    Call COMUNI("Orden de Fabricacion: " & NORFA$ & " Cerrada Por " & CANTIEQUI# & " Equipo/s")
    Label34 = ""
    Hide
99  Screen.MousePointer = 1
    BOTREC.Enabled = True
End Sub

Private Sub Command1_Click()
   Command1.Enabled = False
   TPI = LISNUM.TopIndex
   For U& = 1 To LISNUM.ListCount
     If LISNUM.Selected(U& - 1) = True Then
         LISNUM.Selected(U& - 1) = False
       Else
         LISNUM.Selected(U& - 1) = True
     End If
   Next U&
   On Error Resume Next
   LISNUM.TopIndex = TPI
   On Error GoTo 0
   DoEvents
   Command1.Enabled = True

End Sub

Private Sub Command2_Click()
   LISNUM.Clear
   Unload Me
End Sub


Private Sub Command3_Click()
   If DERACCE%("CAMBIDEP", "CAMBIO DE DEPOSITOS") > 1 Then
    Call SELECHO("TADEPOSI")
    If VALACT1$("TADEPOSI") <> "" Then
      Label3 = TRIM$(VALACT1$("TADEPOSI"))
    End If
   End If
  

End Sub

Private Sub Command6_Click()
   If DERACCE%("CAMBIDEP", "CAMBIO DE DEPOSITOS") > 1 Then
    Call SELECHO("TADEPOSI")
    If VALACT1$("TADEPOSI") <> "" Then
      Label1 = TRIM$(VALACT1$("TADEPOSI"))
    End If
   End If
End Sub

Private Sub Form_Load()
     'Actualizacion de Depositos de consumo donde el mismo este vacio con movimiento 'cf'
     Screen.MousePointer = 11
     SELNUMSE.Visible = True
     UTILIZA_SKIN% = 1
     Call APLICACIONSKINS(Me)

     DepoDes% = XVALO(Control("", "DEPODEST"))
     CONDI$ = "CODIMOVI = 'CF' AND CANTSALID > 0 AND (DEPOMOVI < 1 OR DEPOMOVI IS NULL)"
     Call ACTUREGISTRO("MOVISTO", "DEPOMOVI", CONDI$, DepoDes%, REGAF&, "NOMESS")
     Screen.MousePointer = 1
     '
End Sub

Private Sub label1_Change()
   Label8 = ECOARCH$("TADEPOSI", Chr$(XVALO(Label1)))
End Sub

Private Sub Label3_Change()
   Label2 = ECOARCH$("TADEPOSI", Chr$(XVALO(Label3)))
End Sub

Private Sub Label34_Change()
     '
     If Label34 = "" Then
        Label1 = "": Label3 = ""
        LISNUM.Clear
        Exit Sub
     End If
     '
     NORFA$ = Label34
     CI% = VALOBADA("ORDEFABR", "Cod_Inte", "IdSubOr = '" & NORFA$ & "'", "NOMESS")
     CONDI$ = "IdSubOr = '" & NORFA$ & "' AND Cod_Inte = " & CI% & " AND DEPONUMSE < 0 AND Dic_Ensa LIKE 'A%'"
     Call CARGALISTA(SELNUMSE.LISNUM, "TANUMSE", "NUMEROSERIE/A16; Descrip/A33", CONDI$)
     '
     'CARGA DEPOSITOS SEGUN CONFIGURACION O SI TIENE ASIGNADO ARMADOR EXTERNO
     SELNUMSE.Label1 = Control("", "DEPOCIE")
     SELNUMSE.Label3 = Control("", "DEPODEST")
     If ValidaArmador%(NORFA$) > 0 Then ' CASO ARMADOR EXTERNO
       DEPARMA% = XVALO(Control("", "DEPOARMA"))
       Label3 = DEPARMA%
     End If
     '
     For II2% = 1 To SELNUMSE.LISNUM.ListCount
        SELNUMSE.LISNUM.Selected(II2% - 1) = True
     Next II2%
     '
End Sub

