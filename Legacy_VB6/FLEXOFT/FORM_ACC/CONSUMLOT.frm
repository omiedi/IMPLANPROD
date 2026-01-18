VERSION 5.00
Begin VB.Form CONSUMLOT 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta de Consumos por Lote"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   10710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   10710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture6 
      Height          =   1670
      Left            =   9800
      ScaleHeight     =   1605
      ScaleWidth      =   1035
      TabIndex        =   33
      Top             =   0
      Visible         =   0   'False
      Width           =   1100
      Begin VB.Image Image3 
         Height          =   2000
         Left            =   0
         Picture         =   "CONSUMLOT.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      Left            =   105
      MultiSelect     =   2  'Extended
      TabIndex        =   28
      Top             =   5040
      Visible         =   0   'False
      Width           =   10455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3050
      Index           =   4
      Left            =   7455
      TabIndex        =   21
      Top             =   1890
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3050
      Index           =   5
      Left            =   7900
      TabIndex        =   20
      Top             =   1890
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3050
      Index           =   6
      Left            =   9070
      TabIndex        =   19
      Top             =   1890
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3050
      Index           =   7
      Left            =   10250
      TabIndex        =   18
      Top             =   1890
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3050
      Index           =   3
      Left            =   4500
      TabIndex        =   17
      Top             =   1890
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3050
      Index           =   2
      Left            =   3045
      TabIndex        =   16
      Top             =   1890
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3050
      Index           =   1
      Left            =   1525
      TabIndex        =   15
      Top             =   1890
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3050
      Index           =   0
      Left            =   1080
      TabIndex        =   14
      Top             =   1890
      Width           =   15
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Descripción"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   105
      TabIndex        =   12
      Top             =   630
      Width           =   4800
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
         Height          =   495
         Left            =   210
         TabIndex        =   13
         Top             =   360
         Width           =   4425
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Identificador de Lote"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1545
      Left            =   5040
      TabIndex        =   8
      Top             =   105
      Width           =   4605
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "F.Ven:"
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
         Left            =   2520
         TabIndex        =   32
         Top             =   840
         Width           =   750
      End
      Begin VB.Label Label17 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   3360
         TabIndex        =   31
         Top             =   735
         Width           =   1065
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "L.Prov:"
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
         Left            =   -210
         TabIndex        =   30
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label14 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   735
         TabIndex        =   29
         Top             =   735
         Width           =   1800
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   2940
         TabIndex        =   27
         Top             =   1155
         Width           =   1485
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo:"
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
         Left            =   2100
         TabIndex        =   26
         Top             =   1200
         Width           =   750
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   735
         TabIndex        =   25
         Top             =   1155
         Width           =   1485
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   3360
         TabIndex        =   24
         Top             =   315
         Width           =   1065
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "F.Alta:"
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
         Left            =   2520
         TabIndex        =   23
         Top             =   420
         Width           =   750
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   735
         TabIndex        =   22
         Top             =   315
         Width           =   1800
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Lote:"
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
         Left            =   -210
         TabIndex        =   10
         Top             =   420
         Width           =   855
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cant:"
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
         Left            =   -105
         TabIndex        =   9
         Top             =   1200
         Width           =   750
      End
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
      Height          =   2760
      Left            =   120
      MultiSelect     =   2  'Extended
      TabIndex        =   3
      Top             =   2205
      Width           =   10455
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00404080&
      FillColor       =   &H00400000&
      ForeColor       =   &H00400000&
      Height          =   1670
      Left            =   9800
      ScaleHeight     =   1605
      ScaleWidth      =   2715
      TabIndex        =   0
      Top             =   0
      Width           =   2775
      Begin VB.CommandButton Command1 
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
         Left            =   105
         Picture         =   "CONSUMLOT.frx":2398
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Terminar - Salir de Consulta de Stocks"
         Top             =   210
         Width           =   650
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   620
         Left            =   105
         Picture         =   "CONSUMLOT.frx":24E2
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Imprime Ficha de Movimientos de Material (Ficha Kardex)"
         Top             =   860
         Width           =   650
      End
   End
   Begin VB.Line Line2 
      X1              =   10560
      X2              =   10560
      Y1              =   1920
      Y2              =   4960
   End
   Begin VB.Line Line1 
      X1              =   105
      X2              =   10560
      Y1              =   4960
      Y2              =   4960
   End
   Begin VB.Label Label5 
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
      Left            =   4200
      TabIndex        =   11
      Top             =   210
      Width           =   750
   End
   Begin VB.Label Label4 
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
      Left            =   1050
      TabIndex        =   7
      Top             =   210
      Width           =   2010
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "U-M:"
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
      Left            =   3255
      TabIndex        =   6
      Top             =   255
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Código:"
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
      Left            =   0
      TabIndex        =   5
      Top             =   250
      Width           =   960
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"CONSUMLOT.frx":2B4C
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   105
      TabIndex        =   4
      Top             =   1920
      Width           =   10455
   End
   Begin VB.Menu menu 
      Caption         =   "Menú"
      Begin VB.Menu imprimir 
         Caption         =   "Imprimir"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "CONSUMLOT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Hide
End Sub

Private Sub Command8_Click()
   '
   Command8.Enabled = False
   Call MENSERR(24, "No Disponible en la Presente Versión.")
99 Command8.Enabled = True
   '
End Sub

Private Sub Form_Load()
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture6.Visible = True
      Me.BackColor = &HFCD7B6
      Me.Frame1.BackColor = &HFCD7B6
      Me.Frame2(1).BackColor = &HFCD7B6
      Me.Frame4.BackColor = &HFCD7B6
      Me.Label15.BackColor = &HFCD7B6
      
    End If

End Sub

Private Sub imprimir_Click()
   Call Command8_Click
End Sub

Private Sub Label1_Change()
   Label8 = ""
   Label9 = ""
   Label13 = ""
   Label14 = ""
   Label17 = ""
   List1.Clear: List2.Clear
   '
   IDLOTAB$ = TRIM$(Label1)
   CIXI% = CODINT(Label4)
   '
   If IDLOTAB$ <> "" And CIXI% > 0 Then
     '
     Call ABRESTOCKS
20   SQLX$ = " SELECT count(Idenlote)as CantRegistros FROM LOTINGRE "
     SQLX$ = SQLX$ + "WHERE IdenLote = '" & IDLOTAB$ & "' "
     SQLX$ = SQLX$ + "AND Cod_Inte = " & CIXI% & " "
     Set LotesingTemp = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
     With LotesingTemp
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
          CANREGIS% = !CantRegistros
      End If
     End With
     '
     If CANREGIS% > 1 Then
       SQLX$ = " SELECT * FROM LOTINGRE "
       SQLX$ = SQLX$ + "WHERE IdenLote = '" & IDLOTAB$ & "' "
       SQLX$ = SQLX$ + "AND Cod_Inte = " & CIXI% & " "
       Set ElimloTemp = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
       '
       With ElimloTemp
         On Error Resume Next
         .MoveLast
           If Err < 1 Then
              ElimloTemp.Delete
           End If
       End With
       Call ACSALOTE(CIXI%, IDLOTAB$)
       GoTo 20
     End If
     '
     Call ACSALOTE(CIXI%, IDLOTAB$)
     '
     List1.Clear: List2.Clear
     '
     Call ABRESTOCKS
     SQLX$ = " SELECT * FROM LOTINGRE "
     SQLX$ = SQLX$ + "WHERE IdenLote = '" & IDLOTAB$ & "' "
     SQLX$ = SQLX$ + "AND Cod_Inte = " & CIXI% & " "
     Set LotIngTemp = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
     '
     With LotIngTemp
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         FEXALTA = CVINT%(!FechAlta)
         CANORIG = !CanAlta
         CANSALDOX = !CanSaldo
         '
         Label8 = FECHATEX$(FEXALTA)
         Label9 = FORMATNUM(CANORIG, "F10,1")
         Label13 = FORMATNUM(CANSALDOX, "F10,1")
         SARRAS# = CANORIG
         '
         Call ABRESTOCKS
         SQLX$ = " SELECT * FROM MOVISTO "
         SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CIXI% & " "
         Set MoviTemp = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
         With MoviTemp
           On Error Resume Next
           .MoveFirst
           If Err < 1 Then
             Do While Not .EOF
                 LOTCO$ = !IdenLote
                 If LOTCO$ = IDLOTAB$ Then
                   CACONS# = ROUND(!CANTSALID)
                   'If !CodiMovi = "AJ" Then GoTo 39 ' porque los ajustes están incluidos en la cantidad de origen
                   If !CodiMovi = "TR" Then CACONS# = CACONS# - !CANTINGRE
                   If Abs(CACONS#) >= 0.05 Then
                      TTXX$ = Space$(256)
                      FEX = CVINT(!FECHMOV)
                        Call REPLA(TTXX$, FECHATEX$(FEX), 1, 8)
                        Call REPLA(TTXX$, !CodiMovi, 11, 2)
                        Call REPLA(TTXX$, !DoPriCor, 15, 12)
                        Call REPLA(TTXX$, !DoSecCor, 29, 12)
                        Call REPLA(TTXX$, !ReferCor, 43, 26)
                      DEPOM$ = FORMATNUM$(!DEPOMOVI, "I3")
                        Call REPLA(TTXX$, DEPOM$, 70, 3)
                        Call REPLA(TTXX$, FORMATNUM$(CACONS#, "F10.1"), 75, 10)
                      For KKU& = List2.ListCount To 1 Step -1
                        If FECHANUM(Left$(List2.List(KKU& - 1), 8)) <= FEX Then
                          List2.AddItem TTXX$, KKU&
                          GoTo 39
                        End If
                      Next KKU&
                      List2.AddItem TTXX$, 0
                   End If
                 End If
39           .MoveNext
             Loop
           End If
         End With
         '
         ' RESERVAS
         Call APERBASDAT("RESERVA")
         SQLZ$ = " SELECT * FROM RESERVA "
         SQLZ$ = SQLZ$ + "WHERE Cod_Inte = " & CIXI% & " "
         SQLZ$ = SQLZ$ + "AND Status = 0 "     ' para no considerar los marcados para anular
         SQLZ$ = SQLZ$ + "ORDER BY FechRes ASC"
         Set ReserTemp = Stocks.OpenRecordset(SQLZ$, dbOpenSnapshot)
         '
         With ReserTemp
           On Error Resume Next
           .MoveFirst
           If Err < 1 Then
              On Error GoTo 0
              Do While Not .EOF
                CSALI# = !CantRes - !CantCons
                If CSALI# >= 0.05 Then
                  LCONSU$ = ""
                  On Error Resume Next
                  LCONSU$ = !IdenLote
                  On Error GoTo 0
                  If LCONSU$ = IDLOTAB$ Then
                    CACONS# = ROUND(CSALI#, 1)
                    TTXX$ = Space$(256)
                    FEX = CVINT(!FechRes)
                      Call REPLA(TTXX$, FECHATEX$(FEX), 1, 8)
                      Call REPLA(TTXX$, "RF", 11, 2)
                      Call REPLA(TTXX$, !IdSubOr, 15, 12)
                      Call REPLA(TTXX$, "", 29, 12)
                      Call REPLA(TTXX$, "RESERVA FABRICACION", 43, 26)
                    DEPOM$ = FORMATNUM$(0, "I3")
                      Call REPLA(TTXX$, DEPOM$, 70, 3)
                      Call REPLA(TTXX$, FORMATNUM$(CACONS#, "F10.1"), 75, 10)
                    For KKU& = List2.ListCount To 1 Step -1
                      If FECHANUM(Left$(List2.List(KKU& - 1), 8)) <= FEX Then
                        List2.AddItem TTXX$, KKU&
                        GoTo 49
                      End If
                    Next KKU&
                    List2.AddItem TTXX$, 0
                  End If
                End If
49            .MoveNext
              Loop
           End If
         End With
         '
         SARRAS# = CANORIG
         For KKU& = 1 To List2.ListCount
            TTXX$ = List2.List(KKU& - 1)
            SARRAS# = SARRAS# - XVALO(Mid$(TTXX$, 75, 10))
            Call REPLA(TTXX$, FORMATNUM$(SARRAS#, "F11.1"), 85, 11)
            List1.AddItem TTXX$
         Next KKU&
         '
       End If
     End With
     Label14 = LOTORIG$(Label1.Caption)
     Label17 = FECHATEX$(FECHVENLOT%(Label1.Caption))
     
   End If
   Screen.MousePointer = 1
   '
End Sub


Private Sub Label4_Change()
   Label5 = ""
   Label6 = ""
   '
   CIXI% = CODINT%(Label4)
   If CIXI% > 0 Then
     Label5 = UNIMED$(CIXI%)
     Label6 = DESCRIT0$(CIXI%)
   End If
   '
End Sub

Private Sub salir_Click()
   Call Command1_Click
End Sub
