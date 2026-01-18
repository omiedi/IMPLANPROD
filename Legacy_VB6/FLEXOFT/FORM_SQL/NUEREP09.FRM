VERSION 5.00
Begin VB.Form NUEREP09 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nueva Orden de Reparación"
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8385
   ForeColor       =   &H00404040&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   8385
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text6 
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
      Left            =   4440
      TabIndex        =   26
      Top             =   2640
      Width           =   2895
   End
   Begin VB.TextBox Text5 
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
      Left            =   2280
      TabIndex        =   19
      Top             =   4440
      Width           =   5055
   End
   Begin VB.TextBox Text4 
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
      Left            =   2280
      TabIndex        =   18
      Top             =   4080
      Width           =   5055
   End
   Begin VB.TextBox Text3 
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
      Left            =   2280
      TabIndex        =   17
      Top             =   3720
      Width           =   5055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "."
      Height          =   255
      Left            =   1920
      TabIndex        =   16
      Top             =   3360
      Width           =   135
   End
   Begin VB.TextBox Text2 
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
      Left            =   960
      TabIndex        =   15
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox Text1 
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
      Left            =   2280
      TabIndex        =   6
      Top             =   600
      Width           =   5055
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00400000&
      Height          =   8610
      Left            =   7440
      ScaleHeight     =   8550
      ScaleWidth      =   1320
      TabIndex        =   0
      Top             =   0
      Width           =   1380
      Begin VB.CommandButton Command4 
         Caption         =   "O.K."
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
         Left            =   120
         Picture         =   "NUEREP09.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Aprobar Control de Salida"
         Top             =   6930
         Width           =   690
      End
      Begin VB.CommandButton Command11 
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
         Left            =   120
         Picture         =   "NUEREP09.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Salir"
         Top             =   225
         Width           =   690
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Help"
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
         Left            =   120
         Picture         =   "NUEREP09.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   975
         Width           =   690
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   630
         TabIndex        =   2
         Top             =   7800
         Width           =   690
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -285
         Picture         =   "NUEREP09.frx":075E
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   8040
         Width           =   1515
      End
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Sector:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   3600
      TabIndex        =   27
      Top             =   2640
      Width           =   1245
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Ult.Técnico"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   3120
      TabIndex        =   25
      Top             =   2160
      Width           =   1245
   End
   Begin VB.Label Label13 
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
      Left            =   4440
      TabIndex        =   24
      Top             =   2160
      Width           =   2895
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "F.U.Rep:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   1200
      TabIndex        =   23
      Top             =   2640
      Width           =   1005
   End
   Begin VB.Label Label10 
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
      Left            =   2280
      TabIndex        =   22
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Garantía:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   1200
      TabIndex        =   21
      Top             =   2160
      Width           =   1005
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
      Left            =   2280
      TabIndex        =   20
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
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
      Left            =   6240
      TabIndex        =   14
      Top             =   240
      Width           =   1065
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Talón:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   5400
      TabIndex        =   13
      Top             =   285
      Width           =   1035
   End
   Begin VB.Label Label9 
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
      Left            =   4800
      TabIndex        =   12
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Código:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   3840
      TabIndex        =   11
      Top             =   1100
      Width           =   870
   End
   Begin VB.Label Label7 
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
      Left            =   2280
      TabIndex        =   10
      Top             =   3360
      Width           =   5055
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
      Height          =   525
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Width           =   7230
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   120
      TabIndex        =   8
      Top             =   3360
      Width           =   885
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nº de Serie  >>>>>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   120
      TabIndex        =   7
      Top             =   650
      Width           =   2235
   End
End
Attribute VB_Name = "NUEREP09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command11_Click()
    '
    Unload Me
    '
End Sub

Private Sub Command4_Click()
   '
   Command4.Enabled = False
   '
   REMI$ = TRIM$(Label7)
   NC% = XVALO(Label4)
   '
   'VALIDACIONES
   If TRIM$(REMI$) = "" And NC% < 1 Then
      Call MENSERR(24, "No Hay un Remito Seleccionado")
      GoTo 99
   End If
   '
   If List1.ListCount > 0 Then
      Call MENSERR(24, "Imposible Grabar Control Quedan Bultos sin Controlar")
      GoTo 99
   End If
   '
   If List2.ListCount < 1 Then
      Call MENSERR(24, "Imposible Grabar Control No Hay Bultos Controlados")
      GoTo 99
   End If
   '
   'GENNERO CONDICION Y GRABO
   CONDI$ = " Nume_Clie = " & NC% & " "
   CONDI$ = CONDI$ + " AND CodiMovi = 'RT' "
   CONDI$ = CONDI$ + " AND DoPriLar =  '" & REMI$ & "'"
   CONDI$ = CONDI$ + " AND COD_INTE > 0"
   '
   USERN% = USNBR% Mod 100
   '
   Call ACTUREGISTRO("MOVISTO", "VERIFICADOPOR", CONDI$, USERN%, REGAF&, "NOMESS")
   'ESTE ACTUREGISTRO SE LE HACE UN Left(Now, 19) PARA QUE NO GUARDE EL PREFIJO AM O PM POR QUE SI NO DA ERROR
   Call ACTUREGISTRO("MOVISTO", "FHor_Verif", CONDI$, AHORA$, REGAF&)
   '
   Call COMUNI("Control de Salida Realizado")
   Call BLANFORMU
   Text1 = "" ' caja de texto del codigo de barra del remito
   Text1.SetFocus

99 Command4.Enabled = True

End Sub

Private Sub Form_Load()
   '
   NROREP& = 1 + XVALO(VALOBADA("REPARA", "MAX(NUMREPA)", CONDI$, "NOMESS"))
   Label11 = Str(NROREP&)
   '
End Sub

Private Sub Text1_Change()
    Call BLANFORMU
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
      Exit Sub
    End If
    '
    Call ABRESTOCK
    LECTURA$ = TRIM$(Text1.TEXT)
    '
    CONDI$ = "NUMEROSERIE='" & LECTURA$ & "'"
    CI% = XVALO(VALOBADA("TANUMSE", "COD_INTE", CONDI$))
    CODIGO$ = CODEXT$(CI%)
    DESCRI$ = DESCRIT0$(CI%)
    ENGARAN$ = "NO"
    FE% = HOY(HO$)
    FECHCARANT% = FeVenGarant(NUMESER$)
    If FECHCARANT% > 0 Then
      If FE% < FECHCARANT% Then
        ENGARAN$ = "SI"
      End If
    End If
    ULTIMOTECNICO$ = SAFETEXT$(VALOBADA("TANUMSE", "", CONDI$))
    '
    
End Sub

Private Sub Text2_Change()
      Label9 = ""
      Label10 = ""
      'CI% = CODINT(Mid$(Text2, 8, 5))
      CI% = CODINTBARRA%(TRIM$(Text2))
      '
      Label9 = CODEXT$(CI%)
      If CI% < 1 Then Exit Sub
      Label10 = DESCRIT0$(CI%)
      '
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
      '
      If KeyAscii <> 13 Then Exit Sub
      '
      Dim ESTA As Boolean
      '
      ESTA = False
      '
      'CI% = CODINT(Mid$(Text2, 8, 5))
      CI% = CODINTBARRA%(TRIM$(Text2))
      If CI% < 1 Then
         Call MENSERR(24, "Código No Existente en Detalle o Ya Descargado")
         Text2 = ""
         Exit Sub
      End If
      '
      For KI% = 1 To List1.ListCount
            CILIST% = CODINT%(Mid$(List1.List(KI% - 1), 1, 20))
            If CI% = CILIST% Then
              ESTA = True
              List2.AddItem List1.List(KI% - 1)
              List1.RemoveItem (KI% - 1)
              Text2 = ""
              Exit For
            End If
      Next KI%
      '
      If ESTA = False Then
         Call MENSERR(24, "Código No Existente en Detalle o Ya Descargado")
         Text2 = ""
      End If
      '
End Sub

Sub CARGO_INFO_LIST(CI%, QTIDAD)
    '
    CANTACUMUL = 0
    COD$ = CODEXT$(CI%)
    Call REPLA$(TX$, COD$, 1, 20)
    Call REPLA$(TX$, DESCRIT0$(CI%), 21, 40)
    '
    CAN# = QTIDAD
    CANT_FILAS% = CANBULTOS(CI%, CAN#, "VENTA")
    '
    For I% = 1 To CANT_FILAS% - 1 ' -1 PARA GENERAR EL ULTIMO POR EL RESTO
       CANTI = ROUND(LINDIVEN(CI%), 0)
       CANTACUMUL = CANTACUMUL + CANTI
       CANTIX$ = FORMATNUM$(CANTI, "F14." & CNTDC$(CIXI%))
       Call REPLA$(TX$, CANTIX$, 62, 14)
       List1.AddItem TX$
       '
    Next I%
    '
    'CREO EL ULTIMO POR EL RESTO
    CANTI = QTIDAD - CANTACUMUL
    CANTIX$ = FORMATNUM$(CANTI, "F14." & CNTDC$(CIXI%))
    Call REPLA$(TX$, CANTIX$, 62, 14)
    List1.AddItem TX$

End Sub

Sub BLANFORMU()

    List1.Clear
    List2.Clear
    Label4 = ""
    Label6 = ""
    Label7 = ""
    Label9 = ""
    Label10 = ""
    Text2 = ""
    
End Sub
