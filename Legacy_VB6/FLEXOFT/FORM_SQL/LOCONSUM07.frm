VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form LOCONSUM07 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta de Consumos por Lote"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   105
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   22
      Top             =   735
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3050
      Index           =   4
      Left            =   6510
      TabIndex        =   10
      Top             =   2400
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3050
      Index           =   3
      Left            =   5250
      TabIndex        =   9
      Top             =   2400
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3050
      Index           =   2
      Left            =   3885
      TabIndex        =   8
      Top             =   2400
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3050
      Index           =   1
      Left            =   1900
      TabIndex        =   7
      Top             =   2400
      Width           =   15
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Encabezado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   6690
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción:"
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
         Left            =   0
         TabIndex        =   21
         Top             =   1600
         Width           =   1380
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
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
         Left            =   0
         TabIndex        =   20
         Top             =   1200
         Width           =   1380
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1470
         TabIndex        =   19
         Top             =   1575
         Width           =   4950
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1470
         TabIndex        =   18
         Top             =   1155
         Width           =   2010
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Referencia:"
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
         Left            =   0
         TabIndex        =   15
         Top             =   780
         Width           =   1380
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Comprobante:"
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
         Left            =   0
         TabIndex        =   14
         Top             =   360
         Width           =   1380
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4305
         TabIndex        =   13
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1470
         TabIndex        =   12
         Top             =   315
         Width           =   2010
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5250
         TabIndex        =   11
         Top             =   315
         Width           =   1170
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1470
         TabIndex        =   6
         Top             =   735
         Width           =   4950
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
      Top             =   2640
      Width           =   6690
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00004040&
      FillColor       =   &H00400000&
      ForeColor       =   &H00400000&
      Height          =   5490
      Left            =   6930
      ScaleHeight     =   5430
      ScaleWidth      =   1875
      TabIndex        =   0
      Top             =   0
      Width           =   1935
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   100
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   16
         Top             =   4680
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   17
            Top             =   120
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   620
         Left            =   105
         Picture         =   "LOCONSUM07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Terminar - Salir de Consulta de Stocks"
         Top             =   210
         Width           =   650
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   620
         Left            =   105
         Picture         =   "LOCONSUM07.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Imprime Ficha de Movimientos de Material (Ficha Kardex)"
         Top             =   860
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "LOCONSUM07.frx":07B4
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   4920
         Width           =   1515
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "LOCONSUM07.frx":26D6
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "LOCONSUM07.frx":290A
      TabIndex        =   23
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Line Line2 
      X1              =   6830
      X2              =   6830
      Y1              =   2400
      Y2              =   5415
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   6810
      Y1              =   4650
      Y2              =   4650
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Código                   Id.Lote                              Parcial             Total"
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
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   6705
   End
End
Attribute VB_Name = "LOCONSUM07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Label4 = ""
   Label7 = ""
   Label8 = ""
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
   Call APLICACIONSKINS(Me, Picture3)
End Sub

Private Sub Label4_Change()
   List1.Clear
   Dim TXLIST$(2048)
   CANRE% = 0
   DOPRIM$ = TRIM$(Label4)
   '
   If DOPRIM$ <> "" Then
     List1.Clear
     '
     Call ABRESTOCKS
     SQLX$ = " SELECT COD_INTE, CANTSALID, IdenLote FROM MOVISTO WITH(NOLOCK) "
     SQLX$ = SQLX$ + " WHERE DoPriCor = '" & DOPRIM$ & "' "
     SQLX$ = SQLX$ + " and cantsalid > 0"
     'sql
     'Dim MoviTemp As ADODB.Recordset
     'Set MoviTemp = New ADODB.Recordset
     'MoviTemp.CursorType = adOpenDynamic
     'MoviTemp.LockType = adLockReadOnly
     'MoviTemp.Open FILTSQL$(SQLX$), FLEXCONN, , , adCmdText
     '
     Set MoviTemp = READSET(SQLX$)
     With MoviTemp
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
           CODE$ = CODEXT$(XVALO(!cod_inte))
           locon$ = SAFETEXT$(!idenlote)    ' !LotConsum
           CATO = 0
           '
           IDLO$ = locon$   ' TRIM$(Mid$(LOCON$, 32 * KKI% - 31, 24))
           CAPAR = XVALO(!CANTSALID)   ' XVALO(Mid$(LOCON$, 32 * KKI% - 7, 8))
           CATO = CATO + CAPAR
           CANPA$ = FORMATNUM$(CAPAR, "F12.1")
           CANTO$ = Space$(12)
           '
           TTXX$ = Space$(128)
           Call REPLA(TTXX$, CODE$, 1, 16)
           Call REPLA(TTXX$, IDLO$, 19, 16)
           Call REPLA(TTXX$, CANPA$, 37, 12)
           Call REPLA(TTXX$, CANTO$, 49, 12)
           Call REPLA(TTXX$, CODE$, 101, 16)
           '
           CANRE% = CANRE% + 1
           TXLIST$(CANRE%) = TTXX$
         .MoveNext
         Loop
       End If
     End With
     '
     For KK1% = 1 To CANRE%
       For KK2% = KK1% + 1 To CANRE%
         If TXLIST$(KK1%) > TXLIST$(KK2%) Then
           TTXX$ = TXLIST$(KK1%)
           TXLIST$(KK1%) = TXLIST$(KK2%)
           TXLIST$(KK2%) = TTXX$
         End If
       Next KK2%
     Next KK1%
     '
     SUPAR = 0
     For KK1% = 1 To CANRE%
       SUPAR = SUPAR + XVALO(Mid$(TXLIST(KK1%), 37, 12))
       If Left$(TXLIST$(KK1%), 16) <> Left$(TXLIST$(KK1% + 1), 16) Then
         CANTO$ = FORMATNUM$(SUPAR, "F12.1")
         Mid$(TXLIST$(KK1%), 49, 12) = CANTO$
         SUPAR = 0
       End If
     Next KK1%
     '
     For KK1% = CANRE% To 2 Step -1
       If Left$(TXLIST$(KK1%), 16) = Left$(TXLIST$(KK1% - 1), 16) Then
         Mid$(TXLIST$(KK1%), 1, 16) = Space$(16)
       End If
     Next KK1%
     '
     For KK1% = 1 To CANRE%
       List1.AddItem TXLIST(KK1%)
       If KK1% = CANRE% Or Left$(TXLIST$(KK1% + 1), 16) <> Space$(16) Then
         List1.AddItem String$(128, "-")
       End If
     Next KK1%
     '
   End If
   Screen.MousePointer = 1
   '
End Sub

Private Sub Label7_Change()
   '
   Label8 = DESCRIT0$(CODINT%(Label7))
   '
End Sub

Private Sub List1_DblClick()
    TTXX$ = List1.TEXT
    IDLOT$ = TRIM$(Mid$(TTXX$, 19, 16))
    If IDLOT$ <> "----------------" Then
      CONSUMLOT07.Label4 = TRIM$(Mid$(TTXX$, 101, 16))
      CONSUMLOT07.Label1 = IDLOT$
      CONSUMLOT07.Show 1
    End If
End Sub
