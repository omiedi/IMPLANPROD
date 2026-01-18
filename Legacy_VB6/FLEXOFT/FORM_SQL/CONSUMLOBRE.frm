VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form CONSUMLOT07 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta de Consumos por Lote"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   11775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   105
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   38
      Top             =   735
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3050
      Index           =   8
      Left            =   11330
      TabIndex        =   29
      Top             =   2010
      Width           =   15
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
      Top             =   5460
      Visible         =   0   'False
      Width           =   10455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3050
      Index           =   4
      Left            =   7455
      TabIndex        =   21
      Top             =   2010
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3050
      Index           =   5
      Left            =   7900
      TabIndex        =   20
      Top             =   2010
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3050
      Index           =   6
      Left            =   9070
      TabIndex        =   19
      Top             =   2010
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3050
      Index           =   7
      Left            =   10250
      TabIndex        =   18
      Top             =   2010
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3050
      Index           =   3
      Left            =   4500
      TabIndex        =   17
      Top             =   2010
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3050
      Index           =   2
      Left            =   3045
      TabIndex        =   16
      Top             =   2010
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3050
      Index           =   1
      Left            =   1525
      TabIndex        =   15
      Top             =   2010
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3050
      Index           =   0
      Left            =   1080
      TabIndex        =   14
      Top             =   2010
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
      Height          =   1250
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
         Height          =   735
         Left            =   210
         TabIndex        =   13
         Top             =   360
         Width           =   4425
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Identificador de Partida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1785
      Left            =   5040
      TabIndex        =   8
      Top             =   105
      Width           =   5720
      Begin VB.Label Label22 
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
         Left            =   3990
         TabIndex        =   37
         Top             =   1365
         Width           =   1590
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Destino:"
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
         Left            =   2940
         TabIndex        =   36
         Top             =   1395
         Width           =   960
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Status:"
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
         Left            =   3645
         TabIndex        =   35
         Top             =   720
         Width           =   1035
      End
      Begin VB.Label Label19 
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
         Left            =   4725
         TabIndex        =   34
         Top             =   665
         Width           =   855
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ubicación:"
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
         Left            =   2940
         TabIndex        =   33
         Top             =   1080
         Width           =   960
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
         Left            =   3990
         TabIndex        =   32
         Top             =   1015
         Width           =   1590
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Lote Orig.:"
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
         Left            =   -45
         TabIndex        =   31
         Top             =   720
         Width           =   1110
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
         Left            =   1200
         TabIndex        =   30
         Top             =   660
         Width           =   1590
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
         Left            =   1200
         TabIndex        =   27
         ToolTipText     =   "Doble Click para Ajustar Saldos por Lote Según Ingresos y Consumos"
         Top             =   1365
         Width           =   1590
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
         Left            =   315
         TabIndex        =   26
         Top             =   1425
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
         Left            =   1200
         TabIndex        =   25
         Top             =   1020
         Width           =   1590
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
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
         Left            =   4365
         TabIndex        =   24
         Top             =   315
         Width           =   1200
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Alta:"
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
         Left            =   2970
         TabIndex        =   23
         Top             =   375
         Width           =   1350
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
         Left            =   1200
         TabIndex        =   22
         Top             =   315
         Width           =   1590
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ident.:"
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
         Left            =   -95
         TabIndex        =   10
         Top             =   375
         Width           =   1155
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad:"
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
         Left            =   -45
         TabIndex        =   9
         Top             =   1080
         Width           =   1110
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
      Left            =   105
      MultiSelect     =   2  'Extended
      TabIndex        =   3
      Top             =   2325
      Width           =   11500
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00404080&
      FillColor       =   &H00400000&
      ForeColor       =   &H00400000&
      Height          =   1870
      Left            =   10860
      ScaleHeight     =   1815
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
         Height          =   735
         Left            =   105
         Picture         =   "CONSUMLOBRE.frx":0000
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
         Height          =   735
         Left            =   105
         Picture         =   "CONSUMLOBRE.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Re-Impresión de Etiqueta"
         Top             =   960
         Width           =   650
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "CONSUMLOBRE.frx":07B4
      TabIndex        =   39
      Top             =   420
      Visible         =   0   'False
      Width           =   1485
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   525
      OleObjectBlob   =   "CONSUMLOBRE.frx":082E
      Top             =   420
   End
   Begin VB.Line Line2 
      X1              =   11610
      X2              =   11610
      Y1              =   2040
      Y2              =   5070
   End
   Begin VB.Line Line1 
      X1              =   105
      X2              =   11610
      Y1              =   5085
      Y2              =   5085
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
      Left            =   825
      TabIndex        =   7
      Top             =   210
      Width           =   2850
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
      Left            =   3675
      TabIndex        =   6
      Top             =   255
      Width           =   540
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
      Left            =   -160
      TabIndex        =   5
      Top             =   240
      Width           =   960
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"CONSUMLOBRE.frx":0A62
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
      Top             =   2040
      Width           =   11505
   End
End
Attribute VB_Name = "CONSUMLOT07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub command1_Click()
   Hide
End Sub

Private Sub Command8_Click()
   '
   Command8.Enabled = False
   CIXI% = CODINT%(Label4)
   LOTEX$ = TRIM$(Label1)
   If CIXI% > 0 Then
      If LOTEX$ <> "" Then
        Call PRINSTIKAPRO1(CIXI%, LOTEX$)
      End If
   End If
99 Command8.Enabled = True
   '
End Sub

Private Sub Form_Activate()
   '
   Call Label1_Change

End Sub


Private Sub Form_Load()

   Call APLICACIONSKINS(Me)

End Sub

Private Sub Label1_Change()
   Label8 = ""
   Label9 = ""
   Label19 = ""
   Label13 = ""
   Label14 = ""
   Label17 = ""
   '
   List1.Clear: List2.Clear
   '
   IDLOTAB$ = TRIM$(Label1)
   CIXI% = CODINT(Label4)
   '
   If IDLOTAB$ <> "" And CIXI% > 0 Then
     'Call ACSALOTE(CIXI%, IDLOTAB$)
     '
     List1.Clear: List2.Clear
     '
     'sql
     Call ABRESTOCKS
     SQLX$ = " SELECT * FROM LOTINGRE "
     SQLX$ = SQLX$ + "WHERE IdenLote = '" & IDLOTAB$ & "' "
     SQLX$ = SQLX$ + "AND Cod_Inte = " & CIXI% & " "
     Dim LOTINGTEMP As ADODB.Recordset
     Set LOTINGTEMP = New ADODB.Recordset
     LOTINGTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

     'Set LotIngTemp = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
     '
     With LOTINGTEMP
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         FEXALTA = CVINT%(!FECHALTA)
         CANORIG = !CanAlta
         CANSALDOX = !CANSALDO
         '
         Label8 = FECHATEX$(FEXALTA)
'         Label9 = FORMATNUM(CANORIG, "F10,1")
         Label9 = FORMATNUM$(CANORIG, "F14." & CNTDC$(CIXI%) & "")
         'Label13 = FORMATNUM(CANSALDOX, "F10,1")
         Label13 = FORMATNUM$(CANSALDOX, "F14." & CNTDC$(CIXI%) & "")
         
         SARRAS# = CANORIG
         '
         Call ABRESTOCKS
         SQLX$ = " SELECT * FROM MOVISTO WITH(NOLOCK) "
         SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CIXI% & " "
         SQLX$ = SQLX$ + "AND IdenLote = '" & IDLOTAB$ & "' "
         'SQLX$ = SQLX$ + "ORDER BY FechMov ASC"
         SQLX$ = SQLX$ + "ORDER BY FechMov, FeReMovi ASC "
         '
         Dim MoviTemp As ADODB.Recordset
         Set MoviTemp = New ADODB.Recordset
         MoviTemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

         With MoviTemp
           On Error Resume Next
           .MoveFirst
           If Err < 1 Then
              Do While Not .EOF
                   CANIN# = !CantIngre
                   CACONS# = !cantsalid
                   TTXX$ = Space$(256)
                   FEX = CVINT(!fechMov)
                     Call REPLA(TTXX$, FECHATEX$(FEX), 1, 8)
                     Call REPLA(TTXX$, !CodiMovi, 11, 2)
                     Call REPLA(TTXX$, !DoPriCor, 15, 12)
                     Call REPLA(TTXX$, !DoSecCor, 29, 12)
                     Call REPLA(TTXX$, !ReferCor, 43, 26)
                   DEPOM$ = FORMATNUM$(!DEPOMOVI, "I3")
                     Call REPLA(TTXX$, DEPOM$, 70, 3)
                     'Call REPLA(TTXX$, FORMATNUM$(CANIN#, "F10.1"), 75, 10)
                     Call REPLA(TTXX$, FORMATNUM$(CANIN#, "F10." & CNTDC$(CIXI%) & ""), 75, 10)
                     'Call REPLA(TTXX$, FORMATNUM$(CACONS#, "F10.1"), 86, 10)
                     Call REPLA(TTXX$, FORMATNUM$(CACONS#, "F10." & CNTDC$(CIXI%) & ""), 86, 10)
                   For KKU& = List2.ListCount To 1 Step -1
                     If FECHANUM(Left$(List2.List(KKU& - 1), 8)) <= FEX Then
                       List2.AddItem TTXX$, KKU&
                       GoTo 39
                     End If
                   Next KKU&
                   List2.AddItem TTXX$, 0
39            .MoveNext
              Loop
           End If
           On Error GoTo 0
         End With
         MoviTemp.Close
         Set MoviTemp = Nothing
         '
         ' RESERVAS
         Call APERBASDAT("RESERVA")
         sqlZ$ = " SELECT * FROM RESERVA "
         sqlZ$ = sqlZ$ + "WHERE Cod_Inte = " & CIXI% & " "
         sqlZ$ = sqlZ$ + "AND Status = 0 "     ' para no considerar los marcados para anular
         sqlZ$ = sqlZ$ + "ORDER BY FechRes ASC"
         'Set ReserTemp = Stocks.OpenRecordset(SQLZ$, dbOpenSnapshot)
         '
         Dim ReserTemp As ADODB.Recordset
         Set ReserTemp = New ADODB.Recordset
         ReserTemp.Open FILTSQL$(sqlZ$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

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
                  LCONSU$ = !idenLote
                  On Error GoTo 0
                  If LCONSU$ = IDLOTAB$ Then
                    CACONS# = ROUND(CSALI#, 1)
                    TTXX$ = Space$(256)
                    FEX = CVINT(!FechRes)
                      Call REPLA(TTXX$, FECHATEX$(FEX), 1, 8)
                      Call REPLA(TTXX$, "RF", 11, 2)
                      Call REPLA(TTXX$, !IDSUBOR, 15, 12)
                      Call REPLA(TTXX$, "", 29, 12)
                      Call REPLA(TTXX$, "RESERVA FABRICACION", 43, 26)
                    DEPOM$ = FORMATNUM$(0, "I3")
                      Call REPLA(TTXX$, DEPOM$, 70, 3)
                      'Call REPLA(TTXX$, FORMATNUM$(CACONS#, "F10.1"), 86, 10)
                      Call REPLA(TTXX$, FORMATNUM$(CACONS#, "F10." & CNTDC$(CIXI%) & ""), 86, 10)
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
         ' BALAREQMO
         Call APERBASDAT("STOCKS")
         sqlZ$ = " SELECT * FROM BALAREQMO "
         sqlZ$ = sqlZ$ + "WHERE Cod_Inte = " & CIXI% & " "
         sqlZ$ = sqlZ$ + " AND IdenLote = '" & IDLOTAB$ & "'"
         sqlZ$ = sqlZ$ + " AND Kilo_Asig <> 0"
         sqlZ$ = sqlZ$ + "ORDER BY FechMov ASC"
         '
         Set ReserTemp = New ADODB.Recordset
         ReserTemp.Open FILTSQL$(sqlZ$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

         With ReserTemp
           On Error Resume Next
           .MoveFirst
           If Err < 1 Then
              On Error GoTo 0
              Do While Not .EOF
                 CSALI# = XVALO(!Kilo_Asig) - XVALO(!Kilo_Consu)
                 CACONS# = ROUND(CSALI#, XVALO(CNTDC$(CIXI%)))
                 If CACONS# > 0 Then
                    TTXX$ = Space$(256)
                    FEX = CVINT(!fechMov)
                         Call REPLA(TTXX$, FECHATEX$(FEX), 1, 8)
                         Call REPLA(TTXX$, "RF", 11, 2)
                         Call REPLA(TTXX$, SAFETEXT$(!DESTINO), 15, 12)
                         Call REPLA(TTXX$, "", 29, 12)
                         Call REPLA(TTXX$, SAFETEXT$(!IdenReque), 43, 26)
                    DEPOM$ = FORMATNUM$(0, "I3")
                         Call REPLA(TTXX$, DEPOM$, 70, 3)
                         'Call REPLA(TTXX$, FORMATNUM$(CACONS#, "F10.1"), 86, 10)
                         Call REPLA(TTXX$, FORMATNUM$(CACONS#, "F10." & CNTDC$(CIXI%) & ""), 86, 10)
                    For KKU& = List2.ListCount To 1 Step -1
                       If FECHANUM(Left$(List2.List(KKU& - 1), 8)) <= FEX Then
                         List2.AddItem TTXX$, KKU&
                         GoTo 491
                       End If
                    Next KKU&
                    List2.AddItem TTXX$, 0
                 End If
491            .MoveNext
              Loop
           End If
         End With
         '
         SARRAS# = 0
         For KKU& = 1 To List2.ListCount
            TTXX$ = List2.List(KKU& - 1)
            SARRAS# = SARRAS# + XVALO(Mid$(TTXX$, 75, 10))
            SARRAS# = SARRAS# - XVALO(Mid$(TTXX$, 86, 10))
            'Call REPLA(TTXX$, FORMATNUM$(SARRAS#, "F11.1"), 96, 11)
            Call REPLA(TTXX$, FORMATNUM$(SARRAS#, "F11." & CNTDC$(CIXI%) & ""), 96, 11)
            List1.AddItem TTXX$
         Next KKU&
         LOTINGTEMP.Close
         Set LOTINGTEMP = Nothing
         '
       End If
     End With
     ReserTemp.Close
     Set ReserTemp = Nothing
     
     '
   End If
   
   CIXXI% = XVALO(CODINT(Label4))
   CONDI$ = "COD_INTE = " & CIXXI% & " AND IDENLOTE ='" & TRIM$(Label1) & "'"
   
   Label14 = VALOBADA("LOTINGRE", "LOTECOL_PROV", CONDI$)
   Label19 = VALOBADA("LOTINGRE", "LIBERA", CONDI$)
   Label17 = VALOBADA("LOTINGRE", "UBICFIS", CONDI$)
   Label22 = VALOBADA("LOTINGRE", "Cod_Desti", CONDI$)
   '
   Screen.MousePointer = 1
   '
End Sub

Private Sub Label13_DblClick()
   CIXI% = CODINT%(Label4)
   If CIXI% > 0 Then
     IDLOTX$ = TRIM$(Label1)
     If IDLOTX$ <> "" Then
       Call ACSALOTE(CIXI%, IDLOTX$)
       Call Label1_Change
     End If
   End If
End Sub

Private Sub Label4_Change()
   Label5 = ""
   Label6 = ""
   Label1 = ""
   '
   CIXI% = CODINT%(Label4)
   If CIXI% > 0 Then
     Label5 = UNIMED$(CIXI%)
     Label6 = DESCRIT0$(CIXI%)
   End If
   '
End Sub

Private Sub List1_DblClick()
   LINEA$ = List1.List(List1.ListIndex)
   DOPRI$ = Mid$(LINEA$, 15, 12)
   CONDI$ = "DOPRICOR ='" & DOPRI$ & "' AND CODIMOVI='OF' AND CANTINGRE > 0 "
   CAUSA% = XVALO(VALOBADA("MOVISTO", "CAUSA", CONDI$))
   If CAUSA% > 0 Then
      CAUSATEX$ = TRIM$(ECOARCH$("CAUINGRE", MKI$(CAUSA%)))
      If CAUSATEX$ <> "" Then
         Call COMUNI("CAUSA: " & CAUSATEX$)
      End If
   End If
   '
   'si es BLIPACK MUESTRA LOS CONSUMOS CON GENERADOR DE REPORTES
   If InStr(EMPREAC$, "BLIPACK") > 0 Then
       OPRO$ = TRIM$(Mid$(DOPRI$, 4)) 'TOMO SOLO EL NUMERO SIN EL TEXTO OP-
       CONDI$ = "NUORPROD =  '" & OPRO$ & "'"
       NORFA$ = VALOBADA("ORDEFABR", "IDSUBOR", CONDI$, "NOMESS")
       '
       If TRIM$(NORFA$) = "" Then
          Call MENSERR(24, "Registro No Válido Para Visualizar Consumos")
          Exit Sub
       End If
       '
       If EXISTE%("GENREP07.EXE") > 0 Then
           Call FINAL("? MOVCONSU (" & NORFA$ & ")") ' ABRE LISTADO DE G.R. SEGUN LA O.F.
       End If
       '
   End If
End Sub
