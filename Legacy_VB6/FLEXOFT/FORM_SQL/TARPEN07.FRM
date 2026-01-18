VERSION 5.00
Begin VB.Form TARPEN07 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tareas Pendientes"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   9735
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command25 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   8720
      Picture         =   "TARPEN07.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Listado Tareas Pendientes por Responsable"
      Top             =   5900
      Width           =   630
   End
   Begin VB.Frame Frame28 
      Caption         =   "Detalle de Pendientes por Responsable"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   9315
      Begin VB.Frame Frame22 
         Caption         =   "Comentarios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1890
         Left            =   4500
         TabIndex        =   5
         Top             =   4200
         Width           =   4635
         Begin VB.CommandButton Command1 
            Caption         =   "Registrar Comentario"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   1560
            Width           =   3735
         End
         Begin VB.TextBox Text13 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1140
            Left            =   105
            MaxLength       =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   6
            Top             =   315
            Width           =   4395
         End
      End
      Begin VB.Frame Frame23 
         Caption         =   "Detalle"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1890
         Left            =   120
         TabIndex        =   3
         Top             =   4200
         Width           =   4275
         Begin VB.TextBox Text28 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1380
            Left            =   105
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Top             =   315
            Width           =   4035
         End
      End
      Begin VB.ListBox List3 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3480
         Left            =   120
         TabIndex        =   1
         Top             =   630
         Width           =   9000
      End
      Begin VB.Line Line6 
         X1              =   120
         X2              =   9120
         Y1              =   330
         Y2              =   330
      End
      Begin VB.Label Label26 
         BackColor       =   &H80000016&
         Caption         =   "  Nro.  Fecha    Detalle / Descripcion                                     Respons."
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   150
         TabIndex        =   2
         Top             =   360
         Width           =   8910
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         X1              =   130
         X2              =   130
         Y1              =   350
         Y2              =   630
      End
      Begin VB.Line Line5 
         X1              =   9120
         X2              =   9120
         Y1              =   315
         Y2              =   4080
      End
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
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
      Left            =   8640
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "TARPEN07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FILRESPO%
'
Private Sub Command1_Click()
    Command1.Enabled = False
    '
    NTAR& = Val(Label28)
    If NTAR& > 0 Then
      CONDI$ = "Nume_Tar = " & NTAR&
'      TTXX$ = FILTERGETRECORD$("TARINTER", 1, MKS$(NTAR&))
'      If TRIM$(TTXX$) <> "" Then
'        Call REPLA(TTXX$, Text13, 641, 384)
'        Call FILTERSETRECORD("TARINTER", 1, MKS$(NTAR&), TTXX$)
'      End If
      Call ACTUREGISTRO("TARINTER", "Obse_Tar", CONDI$, Text13, RAFE&, "NOMESS")
    End If
    '
'    Call CIERRARCH("TARINTER")
    Command1.Enabled = True
End Sub

Private Sub Command25_Click()
   FADMIFLEX.Option4(FILRESPO%).Value = True
   Call FADMIFLEX.Command25_Click
End Sub

Private Sub Form_Load()
   Caption = "Tareas pendientes - " + UCase$(USERNAME$)
   Call CARLISTAPEN
End Sub

Sub CARLISTAPEN()
   '
   Dim RETEY$(1024)
   ENCARLISTA% = 1
   FILRESPO% = 0     ' numero de usuario
   USNA$ = UCase$(USERNAME$)
   If InStr(USNA$, "BAUER") > 0 Then FILRESPO% = 0
   If InStr(USNA$, "FLAVIO") > 0 Then FILRESPO% = 1
   If InStr(USNA$, "NAT") > 0 Then FILRESPO% = 3
   If InStr(USNA$, "FRANCIS") > 0 Then FILRESPO% = 2
   If InStr(USNA$, "OMAR") > 0 Then FILRESPO% = 5
   If InStr(USNA$, "EZEQ") > 0 Then FILRESPO% = 4
   '
   TODOS% = 0
'   RFF$ = RECFILT$("TARINTER", 5, Chr$(FILRESPO%))
   HOIX$ = FETEXCOR1$(HOY(HOS$))
   CONDI$ = "(Fech_Cier < '30/NOV/1990' OR Fech_Cier Is Null)"
   CONDI$ = CONDI$ & " AND (Fech_Prog <= '" & HOIX$ & "' OR Fech_Prog IS NULL)"
   If FILRESPO% > 0 Then
      CONDI$ = CONDI$ & " AND ITec_Resp = " & FILRESPO%
   End If
   SQLX$ = "SELECT * FROM TARINTER WITH(NOLOCK) WHERE " & CONDI$
   SQLX$ = SQLX$ & " ORDER BY Nume_Tar ASC"
   '
15 TP1 = List3.TopIndex
   TP2 = TP1
   For KU& = TP1 + 1 To TP1 + 20
      If KU& < List3.ListCount Then
         If List3.Selected(KU&) = True Then
            TP2 = KU&
            Exit For
         End If
      End If
   Next KU&
   '
   List3.Clear
   Dim TASEL As ADODB.Recordset
   Set TASEL = New ADODB.Recordset
   Set TASEL = FLEXCONN.Execute(SQLX$)
   'HOII% = HOY(HOS$)
   With TASEL
     Do While Not .EOF
'   For KKU& = 1 To Len(RFF$) Step 4
'     REG& = CVS(Mid$(RFF$, KKU&, 4))
'     XX$ = REGLEIDO$("TARINTER", REG&)
'     NTAR& = CVS(Left$(XX$, 4))
'     FEX = CVI(Mid$(XX$, 65, 2))
'     FECIE% = CVI(Mid$(XX$, 67, 2))
'     If FECIE% < 1 Then
'        If CVI(Mid$(XX$, 75, 2)) <= HOII% Then ' PARA SUPRIMIR LAS POSTERGADAS
'           DETAX$ = RTrim(Mid$(XX$, 129, 512))
'           If TRIM$(DETAX$) = "" Then
'             DETAX$ = RTrim(Mid$(XX$, 77, 52))
'           End If
'           Call FRATEXTO(DETAX$, 64)
       NTAR& = !Nume_Tar
       FEX = CVINT(!Fech_Emis)
       DETAX$ = RTrim(SAFETEXT$(!Deta_Tar))
       If TRIM$(DETAX$) = "" Then
          DETAX$ = RTrim(SAFETEXT$(!Des_Corta))
       End If
       Call FRATEXTO(DETAX$, 64)
           '
           CARETEY% = CARETEX%
           CARETEZ% = CARETEX%
           For KKI% = 1 To CARETEX%
              RETEY$(KKI%) = RETEX$(KKI%)
              RETEX$(KKI%) = ""
           Next KKI%
           '
'           OBSEX$ = RTrim$(Mid$(XX$, 641, 384))
       OBSEX$ = RTrim$(SAFETEXT$(!Obse_Tar))
           Call FRATEXTO(OBSEX$, 64)
           If CARETEX% > CARETEZ% Then CARETEZ% = CARETEX%
           '
           YY$ = Space$(128 + 64)
           Call REPLA(YY$, FORMATNUM$(NTAR&, "F5.0"), 1, 5)
           Call REPLA(YY$, FECHATEX$(FEX), 9, 8)
           Call REPLA(YY$, RETEY$(1), 19, 64)
           Call REPLA(YY$, RETEX$(1), 97, 64)
           '
           List3.AddItem YY$
           For KKI% = 2 To CARETEZ%
             YY$ = Space$(128 + 64)
             Call REPLA(YY$, RETEY$(KKI%), 19, 64)
             Call REPLA(YY$, RETEX$(KKI%), 97, 64)
             List3.AddItem YY$
           Next KKI%
           List3.AddItem String$(192, "-")
     .MoveNext
     Loop
   End With
'        End If
'     End If
'   Next KKU&
   TASEL.Close
   Set TASEL = Nothing
   '
   If TP1 < 1 Then TP1 = List3.ListCount - 15
   On Error Resume Next
   List3.TopIndex = TP1
   If TP2 > 0 Then List3.Selected(TP2) = True
   On Error GoTo 0
   ENCARLISTA% = 0
   '
End Sub
'
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Cancel = 0
   If UnloadMode = 0 Then
      Cancel = 1
      Me.WindowState = 1    ' LO MINIMIZA
   End If
End Sub

Private Sub Label28_Change()
   Text28.TEXT = ""
   Text13.TEXT = ""
   '
   NTAR& = Val(Label28)
   If NTAR& > 0 Then
     CONDI$ = "Nume_Tar = " & NTAR&
     'TTXX$ = FILTERGETRECORD$("TARINTER", 1, MKS$(NTAR&))
     'FEX = CVI(Mid$(TTXX$, 65, 2))
     FEX = CVINT(VALOBADA("TARINTER", "Fech_Emis", CONDI$, "NOMESS"))
     Label29 = FECHATEX$(FEX)
'     FEX = CVI(Mid$(TTXX$, 67, 2))
     FEX = CVINT(VALOBADA("TARINTER", "Fech_Cier", CONDI$, "NOMESS"))
     Text31 = FECHATEX$(FEX)
'     FEPOST = CVI(Mid$(TTXX$, 75, 2))
     FEPOST = CVINT(VALOBADA("TARINTER", "Fech_Prog", CONDI$, "NOMESS"))
     If FEPOST > HOY(HOS$) Then
        Label34 = FECHATEX$(FEPOST)
     End If
     Text29 = TRIM$(VALOBADA("TARINTER", "Refe_Tar", CONDI$, "NOMESS"))
     ' Text29 = TRIM$(Mid$(TTXX$, 77, 52))
     'Text28 = RTrim(Mid$(TTXX$, 129, 512))
     Text28 = RTrim$(VALOBADA("TARINTER", "Deta_Tar", CONDI$, "NOMESS"))
     'Text13 = RTrim(Mid$(TTXX$, 641, 384))
     Text13 = RTrim$(VALOBADA("TARINTER", "Obse_Tar", CONDI$, "NOMESS"))
     'IRESPO% = Asc(Mid$(TTXX$, 69, 1))
     IRESPO% = XVALO(VALOBADA("TARINTER", "ITec_Resp", CONDI$, "NOMESS"))
   End If
   '
End Sub

Private Sub List3_Click()
   For KU& = List3.ListIndex To 0 Step -1
     NTAR1& = Val(Left$(List3.List(KU&), 5))
     If NTAR1& > 0 Then
       Label28 = TRIM$(Str$(NTAR1&))
       Exit Sub
     End If
   Next KU&
End Sub

Private Sub Text28_GotFocus()
    Text13.SetFocus
End Sub
