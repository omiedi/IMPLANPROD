VERSION 5.00
Begin VB.Form CONPEPEN 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta Rápida de Pedidos Pendientes por Cliente"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   11715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   11715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture6 
      Height          =   930
      Left            =   10100
      ScaleHeight     =   870
      ScaleWidth      =   2715
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   2775
      Begin VB.Image Image7 
         Height          =   930
         Left            =   0
         Picture         =   "CONPEPEN.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   2775
      End
   End
   Begin VB.CommandButton Command2 
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
      Left            =   2100
      TabIndex        =   24
      Top             =   210
      Width           =   150
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Incluir Cumpidos"
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
      Left            =   3465
      TabIndex        =   22
      Top             =   210
      Width           =   1800
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   4365
      Index           =   4
      Left            =   6825
      TabIndex        =   21
      Top             =   1575
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   4365
      Index           =   5
      Left            =   8085
      TabIndex        =   20
      Top             =   1575
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   4365
      Index           =   6
      Left            =   9450
      TabIndex        =   19
      Top             =   1575
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   4365
      Index           =   7
      Left            =   10395
      TabIndex        =   18
      Top             =   1575
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   4365
      Index           =   8
      Left            =   11360
      TabIndex        =   17
      Top             =   1575
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   4365
      Index           =   3
      Left            =   5565
      TabIndex        =   16
      Top             =   1575
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   4365
      Index           =   2
      Left            =   3570
      TabIndex        =   15
      Top             =   1575
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   4365
      Index           =   1
      Left            =   1785
      TabIndex        =   14
      Top             =   1575
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   4365
      Index           =   0
      Left            =   775
      TabIndex        =   13
      Top             =   1575
      Width           =   15
      Begin VB.Line Line3 
         X1              =   0
         X2              =   0
         Y1              =   105
         Y2              =   0
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Rango de Fechas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   5460
      TabIndex        =   6
      Top             =   105
      Width           =   4500
      Begin VB.CommandButton BOTSEL 
         Caption         =   "..."
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
         Index           =   1
         Left            =   4095
         Picture         =   "CONPEPEN.frx":2398
         TabIndex        =   8
         Top             =   420
         Width           =   175
      End
      Begin VB.CommandButton BOTSEL 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   1995
         Picture         =   "CONPEPEN.frx":26A2
         TabIndex        =   7
         Top             =   400
         Width           =   165
      End
      Begin VB.Label Label1 
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
         Height          =   275
         Index           =   1
         Left            =   3150
         TabIndex        =   26
         Top             =   420
         Width           =   960
      End
      Begin VB.Label Label1 
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
         Height          =   275
         Index           =   0
         Left            =   1050
         TabIndex        =   25
         Top             =   420
         Width           =   960
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Desde:"
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
         Left            =   0
         TabIndex        =   10
         Top             =   420
         Width           =   855
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta:"
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
         Left            =   2205
         TabIndex        =   9
         Top             =   420
         Width           =   750
      End
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1155
      TabIndex        =   3
      ToolTipText     =   "Doble Click para Selecionar  Clientes"
      Top             =   210
      Width           =   960
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4380
      Left            =   105
      MultiSelect     =   2  'Extended
      TabIndex        =   2
      ToolTipText     =   "Doble Click para ver Documento Electrónico"
      Top             =   1575
      Width           =   11505
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00800000&
      FillColor       =   &H00400000&
      ForeColor       =   &H00400000&
      Height          =   930
      Left            =   10100
      ScaleHeight     =   870
      ScaleWidth      =   2715
      TabIndex        =   0
      Top             =   0
      Width           =   2775
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   11
         Top             =   105
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   15
            Left            =   0
            TabIndex        =   12
            Top             =   -80
            Width           =   12000
         End
      End
      Begin VB.CommandButton COMMAND1 
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
         Height          =   620
         Index           =   0
         Left            =   880
         Picture         =   "CONPEPEN.frx":29AC
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Salir"
         Top             =   105
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "CONPEPEN.frx":2AF6
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   300
         Width           =   1515
      End
   End
   Begin VB.Label Label6 
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
      Height          =   285
      Left            =   210
      TabIndex        =   23
      Top             =   630
      Width           =   4845
   End
   Begin VB.Line Line2 
      X1              =   11610
      X2              =   11610
      Y1              =   1575
      Y2              =   5985
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      TabIndex        =   5
      Top             =   255
      Width           =   960
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"CONPEPEN.frx":4A18
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   105
      TabIndex        =   4
      Top             =   1260
      Width           =   11505
   End
   Begin VB.Menu menu 
      Caption         =   "Menú"
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "CONPEPEN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BOTSEL_Click(Index As Integer)
    Call CommSel("SELFECHA", BOTSEL(Index), Label1(Index))
End Sub

Private Sub Check1_Click()
   List1.Clear
   Call GENKARDEX
End Sub
Private Sub Command1_Click(Index As Integer)
  CONPEPEN.Hide
End Sub

Private Sub Command2_Click()
  Call CommSel("DEUDOR1", Command2, Text1)
End Sub

Private Sub Form_Load()
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture6.Visible = True
      Me.BackColor = &HFCD7B6
      Me.Frame2(0).BackColor = &HFCD7B6
      Me.Frame4.BackColor = &HFCD7B6
      Me.Check1.BackColor = &HFCD7B6
      '
    End If

   Label1(0) = "01/01/80"
   HOII% = HOY(HOS$)
   Label1(1) = HOS$
End Sub



Private Sub label1_Change(Index As Integer)
  Call CARGAPEPEN

End Sub

Private Sub List1_DblClick()
   TTXX$ = List1.TEXT
   NPX& = XVALO(Mid$(TTXX$, 1, 6))
    If NPX& < 1 Then GoTo 99
    '
    TIDOCUM$ = "Pedido de Cliente"
    TIDOCUX$ = "Pedido ABIERTO"
    TIDOCUY$ = "Pedido Interno"
    NUDOCU$ = TRIM$(Str$(NPX&))
    While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
    NUDOCU$ = NUDOCU$ + " "
    Screen.MousePointer = 11
    '
4   For UI& = ULTREG&("PDOCLST") To 1 Step -1
      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(XX$, TIDOCUM$) > 4 Or InStr(XX$, TIDOCUX$) > 4 Or InStr(XX$, TIDOCUY$) > 4 Then
        PPXX% = InStr(XX$, "Nro.")
        If PPXX% > 0 Then
          XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
        End If
        If InStr(XX$, NUDOCU$) > 4 Then
          DOCUNU& = CVS(Left$(XX$, 4))
          GoTo 5
        End If
      End If
    Next UI&
    '
    Screen.MousePointer = 1
    Call COMUNI("No se Encontró el Documento Impreso\Correspondiente a este Pedido.\  \Si se Trata de un Pedido Pendiente de Entrega\Puede Refrescarlo con la Opción\'Modificación de Pedido en Curso'.")
    On Error Resume Next
    OPPED04.Command46.SetFocus
    Call OPPED04.Command46_Click
    On Error GoTo 0
    
    GoTo 99
    '
5   Call MUESTRADOC(DOCUNU&)
    '
99
   '
End Sub

Private Sub salir_Click()
   Call Command1_Click(0)
End Sub

Private Sub Text1_Change()
   NC% = XVALO(Text1)
   Label6 = RASOCLI$(NC%)
   '
   Call CARGAPEPEN
End Sub

Private Sub Text1_DblClick()
   Call Command2_Click
End Sub





Sub GENKARDEX()
   '
   Static DEPXX%, PRIDEP%
   If PRIDEP% < 1 Then
     Call CARDEPOS
     PRIDEP% = COSU%(1)
   End If
   '
   LU$ = LUDAT$
   If EXISTE%(LU$ + "COKARDEX.RFS") < 1 Then
     Call COPYSTRU("MOVISTO", "COKARDEX")
   End If
   REBLA$ = REGBLAN$("COKARDEX")
   '
   List1.Clear
   CODL$ = Text1
5  CIXX% = CODINT%(CODL$)
   If CIXX% < 1 Then Exit Sub
   '
   DEPXX% = Val(Text4)
   If DEPXX% < 0 Then DEPXX% = 0
   '
   'DESFE% = FECHANUM(Text11)
   'DESDE = CDbl(CVDAT(DESFE%))
   'HASFE% = FECHANUM(Text12)
   'HASTA = CDbl(CVDAT(HASFE%))
   If DESDE > HASTA Then
     Call MENSERR(24, "Rango de Fechas no Válido")
     Exit Sub
   End If
   '
   INCLUVIE% = 0
   If Check1.Value = 1 Then INCLUVIE% = 1
   '
6  CIDD% = CIXX%
   JJ& = 0
   SARRAS# = 0
   INCLUPOS% = 0
   If CIDD% > 0 Then
     'Call BLOQARCH("COKARDEX")
     Call ABRESTOCKS
     '
     If INCLUVIE% = 1 Then
        Call ABRESTOCKVIE
        SQLX$ = " SELECT * FROM MOVISTO "
        SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CIDD% & " "
        'SQLX$ = SQLX$ + "AND Cdbl(FechMov) >= " & DESDE & " "
        SQLX$ = SQLX$ + "AND INT(Cdbl(FechMov)) <= " & HASTA & " "
        If DEPXX% > 0 Then
          SQLX$ = SQLX$ + "AND DepoMovi = " & DEPXX% & " "
        End If
        SQLX$ = SQLX$ + "AND CodiMovi <> 'KU' "
        SQLX$ = SQLX$ + "ORDER BY FechMov ASC, FeReMovi ASC; "
        '
        Set MoviTemp = StockVie.OpenRecordset(SQLX$, dbOpenSnapshot)
        With MoviTemp
          On Error Resume Next
          .MoveFirst
          If Err < 1 Then
            Do While Not .EOF
              If CVINT(!fechMov) < DESFE% Then
                   SARRAS# = SARRAS# + !CantIngre - !Cantsalid
                 Else
                   YY$ = REBLA$
                   Call REPLA(YY$, MKI$(CVINT(!fechMov)), 1, 2)
                   Call REPLA(YY$, MKI$(!Cod_Inte), 3, 2)
                   Call REPLA(YY$, !IdenLote, 5, 16)
                   Call REPLA(YY$, !CodiMovi, 21, 2)
                   Call REPLA(YY$, !DoPriCor, 23, 10)
                   Call REPLA(YY$, !DoSecCor, 33, 10)
                   Call REPLA(YY$, !ReferCor, 43, 30)
                   Call REPLA(YY$, MKD$(!ValoUnit), 73, 8)
                     MONEX$ = "$": If !MoneVal = 2 Then MONEX$ = "U$"
                   Call REPLA(YY$, MONEX$, 81, 2)
                   Call REPLA(YY$, Chr$(!DepoMovi), 83, 1)
                   Call REPLA(YY$, MKI$(!Nume_Clie), 84, 2)
                   Call REPLA(YY$, MKI$(!Nume_Prov), 86, 2)
                   Call REPLA(YY$, MKD$(!CantIngre), 88, 8)
                   Call REPLA(YY$, MKD$(!Cantsalid), 96, 8)
                   Call REPLA(YY$, MKD$(0), 104, 8)
                   Call REPLA(YY$, Mid$(!DoPriCor, 11, 4), 112, 4)
                   Call REPLA(YY$, Mid$(!DoSecCor, 11, 4), 116, 4)
                   Call REPLA(YY$, CVINT(!FeReMovi), 122, 2)
                   Call REPLA(YY$, Chr$(!NumUsuar), 126, 1)
                   JJ& = JJ& + 1
                   'Call GRAREG("!COKARDEX", YY$, JJ&)
                   Call GRAREG("#COKARDEX", YY$, JJ&)
              End If
            .MoveNext
            Loop
          End If
        End With
        '
     End If
     '
     SQLX$ = " SELECT * FROM MOVISTO "
     SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CIDD% & " "
     'SQLX$ = SQLX$ + "AND Cdbl(FechMov) >= " & DESDE & " "
     'SQLX$ = SQLX$ + "AND INT(Cdbl(FechMov)) <= " & HASTA & " "
     If DEPXX% > 0 Then
       If DEPXX% = PRIDEP% Then
           SQLX$ = SQLX$ + "AND (DepoMovi = " & DEPXX% & " OR DepoMovi = 0 OR ISNULL(DepoMovi) = True OR ISEMPTY(DepoMovi) = True) "
         Else
           SQLX$ = SQLX$ + "AND DepoMovi = " & DEPXX% & " "
       End If
     End If
     If INCLUVIE% = 1 Then
       SQLX$ = SQLX$ + "AND CodiMovi <> 'KU' "
     End If
     SQLX$ = SQLX$ + "ORDER BY FechMov ASC, FeReMovi ASC; "
     '
     Set MoviTemp = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
     With MoviTemp
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
           If CVINT(!fechMov) > HASFE% Then
             If INCLUPOS% = 0 Then
               OPXX% = COMALTER%("Se Han Detectado Movimientos POSTERIORES\a la Fecha Límite Seleccionada.\\No Considerarlos\Incluirlos")
               If OPXX% < 2 Then INCLUPOS% = (-1)
               If OPXX% > 1 Then INCLUPOS% = 1
             End If
           End If
           '
           If CVINT(!fechMov) < DESFE% Then
               SARRAS# = SARRAS# + !CantIngre - !Cantsalid
             ElseIf CVINT(!fechMov) <= HASFE% Or INCLUPOS% = 1 Then
               YY$ = REBLA$
               Call REPLA(YY$, MKI$(CVINT(!fechMov)), 1, 2)
               Call REPLA(YY$, MKI$(!Cod_Inte), 3, 2)
               Call REPLA(YY$, !IdenLote, 5, 16)
               Call REPLA(YY$, !CodiMovi, 21, 2)
               Call REPLA(YY$, !DoPriCor, 23, 10)
               Call REPLA(YY$, !DoSecCor, 33, 10)
               Call REPLA(YY$, !ReferCor, 43, 30)
               Call REPLA(YY$, MKD$(!ValoUnit), 73, 8)
                 MONEX$ = "$": If !MoneVal = 2 Then MONEX$ = "U$"
               Call REPLA(YY$, MONEX$, 81, 2)
               Call REPLA(YY$, Chr$(!DepoMovi), 83, 1)
               Call REPLA(YY$, MKI$(!Nume_Clie), 84, 2)
               Call REPLA(YY$, MKI$(!Nume_Prov), 86, 2)
               Call REPLA(YY$, MKD$(!CantIngre), 88, 8)
               Call REPLA(YY$, MKD$(!Cantsalid), 96, 8)
               Call REPLA(YY$, MKD$(0), 104, 8)
               Call REPLA(YY$, Mid$(!DoPriCor, 11, 4), 112, 4)
               Call REPLA(YY$, Mid$(!DoSecCor, 11, 4), 116, 4)
               Call REPLA(YY$, CVINT(!FeReMovi), 122, 2)
               Call REPLA(YY$, Chr$(!NumUsuar), 126, 1)
               JJ& = JJ& + 1
               'Call GRAREG("!COKARDEX", YY$, JJ&)
               Call GRAREG("#COKARDEX", YY$, JJ&)
           End If
         .MoveNext
         Loop
       End If
     End With
     Set MoviTemp = Nothing
     'Movisto.Close
     '
     On Error GoTo 0
     'Call SETULTREG("!COKARDEX", JJ&)
     Call SETULTREG("#COKARDEX", JJ&)
     Screen.MousePointer = 11
     '
     If INCLUVIE% = 1 Then
       Call BLOQARCH("COKARDEX")
       'Call TRACENTRAL("COKARDEX", "")
       Call FILEDUMP("COKARDEX")
       Call FILESORT("COKARDEX", "", 1, 1, "ASC")
       'Call TRALOCAL("COKARDEX", "")
       For KKUL& = 1 To JJ&
         CKX$ = REGLEIDO$("COKARDEX", KKUL&)
         Call GRAREG("#COKARDEX", CKX$, KKUL&)
       Next KKUL&
       Call CIERRARCH("COKARDEX")
     End If
     '
     'FIN& = ULTREG&("!COKARDEX")
     FIN& = ULTREG&("#COKARDEX")
     CASALDO# = 0
     If Abs(SARRAS#) >= 0.05 Then
       TTXX$ = Space$(256)
         FEX = DIANTE(DESFE%)
       Call REPLA(TTXX$, FECHATEX$(FEX), 1, 8)
         REFEM$ = "Saldo Anterior al " + FECHATEX$(DESFE%)
       Call REPLA(TTXX$, REFEM$, 43, 36)
         CASALDO# = SARRAS#
       Call REPLA(TTXX$, FORMATNUM$(CASALDO#, "F12.1"), 95, 12)
       List1.AddItem TTXX$
     End If
     '
     For KU& = 1 To FIN&
       'YY$ = REGLEIDO$("!COKARDEX", KU&)
       YY$ = REGLEIDO$("#COKARDEX", KU&)
       TTXX$ = Space$(256)
         FEX = CVI(Left$(YY$, 2))
       Call REPLA(TTXX$, FECHATEX$(FEX), 1, 8)
         CMOVI$ = Mid$(YY$, 21, 2)
       Call REPLA(TTXX$, CMOVI$, 11, 2)
         DOPRI$ = Mid$(YY$, 23, 10) + Mid$(YY$, 112, 4)
       Call REPLA(TTXX$, DOPRI$, 15, 12)
         DOSEC$ = Mid$(YY$, 33, 10) + Mid$(YY$, 116, 4)
       Call REPLA(TTXX$, DOSEC$, 29, 12)
         REFEM$ = Mid$(YY$, 43, 30)
       Call REPLA(TTXX$, REFEM$, 43, 26)
         DEPOM$ = FORMATNUM$(Asc(Mid$(YY$, 83, 1)), "I3")
       Call REPLA(TTXX$, DEPOM$, 70, 3)
         CAINGRE# = CVD(Mid$(YY$, 88, 8))
       Call REPLA(TTXX$, FORMATNUM$(CAINGRE#, "F10.1"), 75, 10)
         CASALI# = CVD(Mid$(YY$, 96, 8))
       Call REPLA(TTXX$, FORMATNUM$(CASALI#, "F10.1"), 85, 10)
         CASALDO# = CASALDO# + CAINGRE# - CASALI#
       Call REPLA(TTXX$, FORMATNUM$(CASALDO#, "F12.1"), 95, 12)
         IDLOT$ = Mid$(YY$, 5, 16)
       Call REPLA(TTXX$, IDLOT$, 129, 16)
       List1.AddItem TTXX$
     Next KU&
     Call FREEMEM("COKARDEX")
     Screen.MousePointer = 1
     '
 Exit Sub
     Call SETULTREG("COKARDEX", JJ&)
     Call CIERRARCH("COKARDEX")
     Screen.MousePointer = 1
     Call CONECRUN("*COKARDEX/" + CODD$, "Ficha Kardex Indirecta") ' + "/WAIT")
     Screen.MousePointer = 1
    Else
     Call MENSERR(24, "Codigo: " + TRIM$(CODD$) + "\sin Movimientos Según Condición SQL.")
   End If
   '
   'If TRIM$(CODL$) = "" Then GoTo 5
   Screen.MousePointer = 1
   GoTo 5
   '
End Sub

Sub CARGAPEPEN()
    NCL1% = XVALO(Text1)
    List1.Clear
   DESFE% = FECHANUM(Label1(0))
   DESDE = CDbl(CVDAT(DESFE%))
   HASFE% = FECHANUM(Label1(1))
   HASTA = CDbl(CVDAT(HASFE%))
   If DESDE > HASTA Then
     Call MENSERR(24, "Rango de Fechas no Válido")
     Exit Sub
   End If

    
    
    Call ABREPEDCLI
    SQLX$ = " SELECT * "
    SQLX$ = SQLX$ + " FROM PEDDETA "
    SQLX$ = SQLX$ + " WHERE NroClie = " & NCL1% & " "
    SQLX$ = SQLX$ + " AND Status < 8 "
    SQLX$ = SQLX$ + " AND (((CanPed - CantEnt) >= 0.05) OR (TipoPed = 'ABIERTO' AND CDBL(FinVigen) >= " & Int(CDbl(Now)) & ")) "
    SQLX$ = SQLX$ + " AND INT(CDBL(FeSolEnt))  >= " & DESDE & ""
    SQLX$ = SQLX$ + " AND INT(CDBL(FeSolEnt))  <= " & HASTA & ""
    SQLX$ = SQLX$ + " ORDER BY NroPed ASC,FeSolEnt ASC, CodiExt ASC; "
    '
    Set PEDTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With PEDTEMP
      On Error Resume Next
      .Move first
      If Err < 1 Then
        Do While (.EOF = False)
        Call REPLA(TXT$, FORMATNUM$(!NroPed, "F6.0"), 1, 6)
        FEXX$ = FECHATEX$(CVINT(!FeSolEnt))
        Call REPLA(TXT$, FEXX$, 8, 8)
        Call REPLA(TXT$, CODEXT$(!CodiInt), 17, 16)
        Call REPLA(TXT$, DESCRIT$(!CodiInt), 34, 16)
        Call REPLA(TXT$, !CBase, 53, 12)
        Call REPLA(TXT$, !TTalle, 65, 12)
        Call REPLA(TXT$, !TColor, 77, 12)
        Call REPLA(TXT$, FORMATNUM$(!CANPED, "F9.1"), 89, 9)
        SALDO = (!CANPED - !CantEnt)
        If SALDO < 0 Then SALDO = 0
        Call REPLA(TXT$, FORMATNUM$(SALDO, "F9.1"), 98, 9)
        List1.AddItem TXT$
        TXT$ = ""
        .MoveNext
        Loop
      End If
      '
    End With
End Sub
