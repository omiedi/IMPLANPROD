VERSION 5.00
Begin VB.Form SEGUIMIENTOMDB 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SEGUIMIENTO"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10770
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   10770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text12 
      Appearance      =   0  'Flat
      Height          =   900
      Left            =   6600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   24
      Top             =   180
      Width           =   4095
   End
   Begin VB.Frame Frame5 
      Caption         =   "Seguimiento de Consultas"
      Height          =   975
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   6375
      Begin VB.TextBox Text14 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   26
         Text            =   "Código Cliente"
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox Text13 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Height          =   285
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "Teléfono"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "Consulta Nº"
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Height          =   285
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "Razon Social"
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "Código Cliente"
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "Código Cliente"
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "Código Cliente"
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "Cuenta"
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Anotaciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   10620
      Begin VB.TextBox Text11 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3855
         Left            =   8040
         TabIndex        =   23
         Text            =   "Text10"
         Top             =   840
         Width           =   90
      End
      Begin VB.TextBox Text10 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3855
         Left            =   1080
         TabIndex        =   22
         Text            =   "Text10"
         Top             =   840
         Width           =   90
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3885
         Left            =   0
         TabIndex        =   12
         Top             =   840
         Width           =   10575
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   10515
         TabIndex        =   8
         Top             =   315
         Width           =   10580
         Begin VB.Label Label2 
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Fecha     Detalle                                                   Usuario"
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Width           =   10560
         End
         Begin VB.Line Line8 
            X1              =   960
            X2              =   1080
            Y1              =   0
            Y2              =   360
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Frame7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4275
         Left            =   10320
         TabIndex        =   7
         Top             =   630
         Width           =   15
      End
      Begin VB.Line Line1 
         X1              =   10580
         X2              =   10580
         Y1              =   320
         Y2              =   6225
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Novedades"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1950
      Left            =   120
      TabIndex        =   0
      Top             =   6240
      Width           =   10620
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   840
         Width           =   2000
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   50
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "O.K"
         Height          =   435
         Left            =   8400
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1320
         Width           =   750
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1175
         Left            =   1155
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   735
         Width           =   7140
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   425
         Left            =   0
         ScaleHeight     =   360
         ScaleWidth      =   10515
         TabIndex        =   2
         Top             =   315
         Width           =   10580
         Begin VB.Label Label4 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Fecha     Detalle                                                   Usuario"
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   0
            TabIndex        =   5
            Top             =   0
            Width           =   10560
         End
         Begin VB.Line Line2 
            X1              =   0
            X2              =   1155
            Y1              =   420
            Y2              =   420
         End
         Begin VB.Label Label5 
            Caption         =   "Label5"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   15
            Left            =   0
            TabIndex        =   3
            Top             =   420
            Width           =   1380
         End
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Registrar Anotación"
         Height          =   540
         Left            =   9345
         TabIndex        =   11
         Top             =   1365
         Width           =   960
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   10665
         Y1              =   1900
         Y2              =   1900
      End
      Begin VB.Line Line7 
         X1              =   10580
         X2              =   10580
         Y1              =   315
         Y2              =   1900
      End
      Begin VB.Line Line6 
         X1              =   8280
         X2              =   10590
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line5 
         X1              =   8295
         X2              =   8295
         Y1              =   735
         Y2              =   1995
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
         Height          =   690
         Left            =   0
         TabIndex        =   4
         Top             =   1200
         Width           =   1170
      End
   End
End
Attribute VB_Name = "SEGUIMIENTOMDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub Command3_Click()
   Command3.Enabled = False
   NUCOTE$ = TRIM$(Text5)
   If NUCOTE$ = "" Then GoTo 99
   '
   'SI NO EXISTE EL REGISTRO LO CREA
   CONDI$ = "Nume_Cotel = '" & NUCOTE$ & "'"
   If CantRegistros("SEGUIMIENTO", CONDI$, "NOMESS") < 1 Then
      Call CreaRegistro("SEGUIMIENTO", "NUME_COTEL", NUCOTE$, "NOMESS")
   End If
   '
   If TRIM$(Text1) <> "" Then
       Call FRATEXTO(Text1, 56)
       HOO% = HOY(HOS$)
'       '(1)SUPERVISOR
'       USUARIOCOMP$ = "(" & TRIM$(Str$(NUMEUSER%)) & ")" & Left(NOMBREUSUARIO$, 24 - Len(TRIM$(Str$(NUMEUSER%))))

       For KKI% = 1 To CARETEX%
         TTX1$ = Space$(76)
         Call REPLA(TTX1$, RETEX$(KKI%), 11, 56)
         If KKI% = 1 Then
           Call REPLA(TTX1$, Text8, 1, 8)
           Call REPLA(TTX1$, Text9, 69, 24)
         End If
         'Call REGAPP("ANOTAPED", Left$(MKS$(NPX&) + TTXX$, 96))
         TTXX$ = TTXX$ + TTX1$ & vbCrLf
       Next KKI%
       '
       CONDI$ = "Nume_Cotel = '" & NUCOTE$ & "'"
       ANOTA$ = VALOBADA("SEGUIMIENTO", "ANOTADOR", CONDI$, "NOMESS")
       ANOTA$ = ANOTA$ + TTXX$
       ANOTA$ = ANOTA$ + String$(88, "=") & vbCrLf
       '
       Call ACTUREGISTRO("SEGUIMIENTO", "ANOTADOR", CONDI$, ANOTA$, RAFE&, "NOMESS")
'       '\\\OT-08-0098-12/03/08-///
'       SQLX$ = "SELECT HISTORIAL FROM DEUDOR12"
'       SQLX$ = SQLX$ + " WHERE  Nume_Cli = " & NC%
'       Set ANOTMP = New ADODB.Recordset
'       ANOTMP.Open (SQLX$), BASSSEX, adOpenKeyset, adLockPessimistic, adCmdText
'       With ANOTMP
'          If Not (.BOF And .EOF) Then
'             ANOTA$ = SAFETEXT$(!HISTORIAL)
'             ANOTA$ = ANOTA$ + TTXX$
'             ANOTA$ = ANOTA$ + String$(88, "=") & vbCrLf
'             !HISTORIAL = ANOTA$
'             .Update
'           End If
'       End With
'       ANOTMP.Close
'       Set ANOTMP = Nothing
       '
       Text1 = ""
       '
       List1.Clear
       Text5 = ""
       Text5.Refresh
       Text5 = NUCOTE$
       Text5.Refresh
       '
     End If
   
99 Command3.Enabled = True
End Sub

'Private Sub Command4_Click()
'    Command4.Enabled = False
'    '
'    'MUESTRA TODOS LOS CLIENTES TAMBIEN LOS DE BAJA
'    SQLX$ = "SELECT * FROM DEUDOR12 "
'    SQLX$ = SQLX$ + " WHERE Nume_Cli > 0"
''    SQLX$ = SQLX$ + " AND (Stat_Cli >= 0 OR Stat_Cli IS NULL)"
'    SQLX$ = SQLX$ + " ORDER BY Nume_Cli"
'    '
'    Set MENSATMP = BaDaABRE_ADMIN_ALQ.OpenRecordset(SQLX$, dbOpenSnapshot)
'    With MENSATMP
'      Do While Not .EOF
'       COCU$ = XVALO(!NUME_CLI)
'       DESCRI$ = SAFETEXT$(!Raso_Cli)
'       Call REPLA(TTXX$, COCU$, 1, 10)
'       Call REPLA(TTXX$, DESCRI$, 12, 40)
'       JJ& = JJ& + 1
'       FRMBUSCAR.List1.AddItem TTXX$
'       .MoveNext
'      Loop
'    End With
'
'    MENSATMP.Close
'    Set MENSATMP = Nothing
'    '
'    FRMBUSCAR.Caption = "HISTORIAL - CLIENTES"
'    FRMBUSCAR.Show
'    NC% = XVALO(Label7)
'    If NC% < 1 Then GoTo 99
'    '
'    Label7 = TRIM$(Str$(NC%))
'    '
'99  Command4.Enabled = True
'End Sub



Private Sub Command9_Click()
   Hide
End Sub





Private Sub Form_Load()
    Call PENDIENTES.Aplicar_skin(Me)
    HOO% = HOY(HOS$)
    Text8 = HOS$
    Text9 = USERNAME$
End Sub

Private Sub Text5_Change()
   List1.Clear
   NUCOTE$ = TRIM$(Text5)
   If NUCOTE$ = "" Then Exit Sub
   '
   CONDI$ = "Nume_Cotel = '" & NUCOTE$ & "'"
   Text3 = VALOBADA("COTELEF", "NROCLIE", CONDI$, "NOMESS")
   NC% = XVALO(Text3)
   Text4 = RASOCLI$(NC%)
   Text14 = TEL1CLI$(NC%)
   Text12 = VALOBADA("COTELEF", "DetaConsu", CONDI$, "NOMESS")


   XX$ = VALOBADA("SEGUIMIENTO", "ANOTADOR", CONDI$, "NOMESS")
   
   LONGI& = Len(XX$)
   INICIO& = 1
   Do While INICIO& < LONGI&
     FIN& = InStr(INICIO&, XX$, Chr$(10))
     If FIN& > 0 Then
       TTXX$ = Mid$(XX$, INICIO&, FIN& - INICIO&)
       List1.AddItem TTXX$
       
       INICIO& = FIN& + 1
     Else
       TTXX$ = Mid$(XX$, INICIO&)
       List1.AddItem TTXX$
       Exit Do
     End If
   Loop
   '
   ' ESTO ES PARA POSICIONARSE AL FINAL DE LA LISTA
   TPXI& = List1.ListCount - 15
   If TPXI& < 0 Then TPXI& = 0
   List1.TopIndex = TPXI&
   '
   List1.Refresh
   Openforms = DoEvents
   '
End Sub


Private Sub Label11_Click()

End Sub

