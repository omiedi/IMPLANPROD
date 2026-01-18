VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form SALOTES09 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta de Saldos por Lote"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   2400
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   28
      Top             =   600
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3250
      Index           =   5
      Left            =   7680
      TabIndex        =   19
      Top             =   1785
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3250
      Index           =   4
      Left            =   6840
      TabIndex        =   18
      Top             =   1785
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3250
      Index           =   1
      Left            =   3045
      TabIndex        =   11
      Top             =   1785
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3250
      Index           =   3
      Left            =   6000
      TabIndex        =   10
      Top             =   1785
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3250
      Index           =   2
      Left            =   4180
      TabIndex        =   9
      Top             =   1785
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3250
      Index           =   0
      Left            =   1080
      TabIndex        =   8
      Top             =   1785
      Width           =   15
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2910
      Left            =   105
      MultiSelect     =   2  'Extended
      TabIndex        =   3
      Top             =   2100
      Width           =   7650
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      FillColor       =   &H00400000&
      ForeColor       =   &H80000008&
      Height          =   1600
      Left            =   7035
      ScaleHeight     =   1575
      ScaleWidth      =   2640
      TabIndex        =   0
      Top             =   -30
      Width           =   2670
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
         Picture         =   "SALOTES09.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Terminar - Salir de Consulta de Stocks"
         Top             =   240
         Width           =   650
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Print"
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
         Picture         =   "SALOTES09.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Imprime Ficha de Movimientos de Material (Ficha Kardex)"
         Top             =   840
         Width           =   650
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   4080
      OleObjectBlob   =   "SALOTES09.frx":07B4
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   3840
      OleObjectBlob   =   "SALOTES09.frx":09E8
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Rechazado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   26
      Top             =   6000
      Width           =   2445
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cuarentena"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   360
      TabIndex        =   25
      Top             =   5620
      Width           =   2325
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total Aprobado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   480
      TabIndex        =   24
      Top             =   5250
      Width           =   2205
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   2880
      TabIndex        =   23
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   2880
      TabIndex        =   22
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   2880
      TabIndex        =   21
      Top             =   5880
      Width           =   1335
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1050
      TabIndex        =   20
      Top             =   580
      Width           =   5790
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   1680
      TabIndex        =   17
      Top             =   1170
      Width           =   5160
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Depósito"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   -60
      TabIndex        =   16
      Top             =   1250
      Width           =   960
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1050
      TabIndex        =   15
      Top             =   1170
      Width           =   540
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Total General"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6120
      TabIndex        =   14
      Top             =   5970
      Width           =   1485
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   4335
      TabIndex        =   13
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   1050
      TabIndex        =   12
      Top             =   210
      Width           =   2850
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   6090
      TabIndex        =   7
      Top             =   210
      Width           =   750
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "U-M"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5125
      TabIndex        =   6
      Top             =   300
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   -60
      TabIndex        =   5
      Top             =   300
      Width           =   960
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Fecha           Id.Lote                               F.Vencim                                  Saldo     Status    Ubic"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   105
      TabIndex        =   4
      Top             =   1815
      Width           =   7650
   End
End
Attribute VB_Name = "SALOTES09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Cancel = 0
   List1.Clear
   Unload Me
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   Call MENSERR(24, "No Disponible en la Presente Versión.")
   Command8.Enabled = True
End Sub

Private Sub Form_Load()
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 0
End Sub


Private Sub Label4_Change()
    CIXI% = CODINT%(Label4)
    Label5 = Unimed$(CIXI%)
    Label6 = DESCRIT0$(CIXI%)
End Sub

Sub MUESTRASALOT()
   List1.Clear
   '
   CIXI% = CODINT%(Label4)
   DEPO% = XVALO(Label9)
   CADECI% = XVALO(CNTDC$(CIXI%))
   '
   For U& = 1 To ULTREG&("!LILOTSE")
       TTYY$ = REGLEIDO$("!LILOTSE", U&)
       TTXX$ = Space$(128)
         FEALTA$ = Mid$(TTYY$, 17, 8)
       Call REPLA(TTXX$, FEALTA$, 1, 8)
         IDLO$ = Mid$(TTYY$, 1, 16)
       Call REPLA(TTXX$, IDLO$, 11, 16)
         FEVEX$ = Mid$(TTYY$, 47, 8)
       Call REPLA(TTXX$, FEVEX$, 30, 8)
         CONDI$ = "IDENLOTE = '" & IDLO$ & "' AND Cod_Inte = " & CIXI%
         UBIC$ = VALOBADA("LOTINGRE", "UbicFis", CONDI$, " NOMESS")
         STAT$ = VALOBADA("LOTINGRE", "LIBERA", CONDI$, "NOMESS")
         CANSA = XVALO(VALOBADA("LOTINGRE", "(CANALTA - CANCONSU)", CONDI$, "NOMESS"))
       Call REPLA(TTXX$, FORMATNUM$(CANSA, "F12." & CADECI%), 44, 12)
         SUMASAL = SUMASAL + CANSA
         Select Case STAT$
      Case "APRO"
         SUMAPRO = SUMAPRO + CANSA
      Case "RECH"
        SUMARECH = SUMARECH + CANSA
      Case "CUAR"
        SUMACUAR = SUMACUAR + CANSA
   End Select
       Call REPLA(TTXX$, STAT$, 57, 9)
       Call REPLA(TTXX$, UBIC$, 65, 10)
       List1.AddItem TTXX$
   Next U&
   '
   Label7 = TRIM$(FORMATNUM$(SUMASAL, "F12." & CADECI%))
   Label13 = TRIM$(FORMATNUM$(SUMAPRO, "F12." & CADECI%))
   Label12 = TRIM$(FORMATNUM$(SUMACUAR, "F12." & CADECI%))
   Label1 = TRIM$(FORMATNUM$(SUMARECH, "F12." & CADECI%))
   Screen.MousePointer = 1
   '

End Sub
'
Private Sub Label9_Change()
    If XVALO(Label9) > 0 Then
         Label11 = ECOARCH$("TADEPOSI", MKI$(XVALO(Label9)))
      Else
         Label11 = "Todos los Depositos"
    End If
End Sub

Private Sub List1_DblClick()
   '
   TTXX$ = List1.TEXT
   CANTING = XVALO(Mid$(TTXX$, 27, 12))
   '
   If CANTING > 0 Then
      IDLOT$ = TRIM$(Mid$(TTXX$, 11, 16))
      If IDLOT$ <> "" Then
        CONSUMLOT07.Label4 = Label4
        CONSUMLOT07.Label1 = IDLOT$
        CONSUMLOT07.Show 1
      End If
   End If
   '
End Sub
