VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form OPPROMRP07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración de Funcionamiento - MRP-II"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2730
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   35
      Top             =   630
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Filtro por Clientes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   3675
      TabIndex        =   32
      Top             =   3675
      Width           =   3165
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Según Pedidos en Curso"
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
         Left            =   315
         TabIndex        =   33
         Top             =   360
         Width           =   2640
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00006000&
      Height          =   5000
      Left            =   7050
      ScaleHeight     =   4935
      ScaleWidth      =   1005
      TabIndex        =   27
      Top             =   -105
      Width           =   1065
      Begin VB.CommandButton Command3 
         Caption         =   "OF's"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   105
         TabIndex        =   37
         Top             =   2880
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command5 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   105
         Picture         =   "OPPROMRP07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Depuracion del Programa de Entregas"
         Top             =   1920
         Width           =   650
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   105
         Picture         =   "OPPROMRP07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Guarda Parámetros"
         Top             =   1050
         Width           =   650
      End
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
         Height          =   700
         Left            =   105
         Picture         =   "OPPROMRP07.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Abandona Configuración"
         Top             =   350
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   28
         Top             =   3780
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   29
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "OPPROMRP07.frx":075E
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   4080
         Width           =   1515
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Programación Semanal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   3675
      TabIndex        =   24
      Top             =   2520
      Width           =   3165
      Begin VB.OptionButton Option7 
         BackColor       =   &H00C0FFC0&
         Caption         =   "1ra Semana Incompleta = 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   210
         TabIndex        =   26
         Top             =   600
         Width           =   2850
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00C0FFC0&
         Caption         =   "1ra.Semana Incompleta = 53"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   210
         TabIndex        =   25
         Top             =   300
         Value           =   -1  'True
         Width           =   2850
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Deposito Stocks"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   3675
      TabIndex        =   20
      Top             =   1575
      Width           =   3165
      Begin VB.CommandButton Command9 
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
         Left            =   630
         TabIndex        =   22
         Top             =   370
         Width           =   175
      End
      Begin VB.TextBox Text7 
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
         Left            =   210
         TabIndex        =   21
         Text            =   " "
         Top             =   370
         Width           =   435
      End
      Begin VB.Label Label1 
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
         Left            =   945
         TabIndex        =   23
         Top             =   370
         Width           =   2115
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Lotes por Omisión"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   210
      TabIndex        =   13
      Top             =   2205
      Width           =   3270
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         Height          =   1905
         Left            =   2280
         TabIndex        =   18
         Top             =   240
         Width           =   850
         Begin VB.TextBox Text2 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   50
            TabIndex        =   3
            Text            =   " "
            Top             =   195
            Width           =   750
         End
         Begin VB.TextBox Text3 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   50
            TabIndex        =   4
            Text            =   " "
            Top             =   615
            Width           =   750
         End
         Begin VB.TextBox Text4 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   50
            TabIndex        =   5
            Text            =   " "
            Top             =   1035
            Width           =   750
         End
         Begin VB.TextBox Text5 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   50
            TabIndex        =   6
            Text            =   " "
            Top             =   1455
            Width           =   750
         End
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Productos Terminados"
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
         Left            =   -105
         TabIndex        =   17
         Top             =   525
         Width           =   2325
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Subconjuntos"
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
         Left            =   -105
         TabIndex        =   16
         Top             =   945
         Width           =   2325
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Mecanizados"
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
         Left            =   -105
         TabIndex        =   15
         Top             =   1365
         Width           =   2325
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Elementos Comprados"
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
         Left            =   -105
         TabIndex        =   14
         Top             =   1785
         Width           =   2325
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Frecuencia - Plazo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1905
      Left            =   210
      TabIndex        =   11
      Top             =   210
      Width           =   3270
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Período"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   210
         TabIndex        =   19
         Top             =   945
         Width           =   2850
         Begin VB.OptionButton Option5 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Mensual"
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
            Left            =   1575
            TabIndex        =   2
            Top             =   315
            Width           =   1170
         End
         Begin VB.OptionButton Option4 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Semanal"
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
            Left            =   210
            TabIndex        =   1
            Top             =   315
            Width           =   1695
         End
      End
      Begin VB.TextBox Text1 
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
         Left            =   2520
         TabIndex        =   0
         Text            =   " "
         Top             =   525
         Width           =   540
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Plazo de Disposición:"
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
         Left            =   -210
         TabIndex        =   12
         Top             =   525
         Width           =   2640
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Modalidad"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1290
      Left            =   3675
      TabIndex        =   10
      Top             =   210
      Width           =   3165
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Cero Stocks y Mínimos"
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
         Left            =   525
         TabIndex        =   8
         Top             =   600
         Width           =   2430
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "No Considera Negativos"
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
         Left            =   525
         TabIndex        =   9
         Top             =   900
         Width           =   2535
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Proceso Normal"
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
         Left            =   525
         TabIndex        =   7
         Top             =   300
         Value           =   -1  'True
         Width           =   1905
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "OPPROMRP07.frx":2680
      Top             =   1800
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   300
      OleObjectBlob   =   "OPPROMRP07.frx":28B4
      TabIndex        =   36
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "OPPROMRP07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub


Private Sub Command2_Click()
    '
    Dim LOSTN(4), LOST$(4)
    '
    SEMANAL% = 0: PERIO$ = "MENSUAL"
    If Option4.Value = True Then
        SEMANAL% = 1: PERIO$ = "SEMANAL"
    End If
    Call GRACONTROL("CANECOM", "PERIODO", PERIO$)
    '
    PLADISTA% = Val(Text1.TEXT)
    Call GRACONTROL("CANECOM", "PLADISTA", TRIM$(Str$(PLADISTA%)))
    '
    LOST$(3) = TRIM$(Text2.TEXT): LOSTN(3) = Val(LOST$(3))
    LOST$(2) = TRIM$(Text3.TEXT): LOSTN(2) = Val(LOST$(2))
    LOST$(1) = TRIM$(Text4.TEXT): LOSTN(1) = Val(LOST$(1))
    LOST$(0) = TRIM$(Text5.TEXT): LOSTN(0) = Val(LOST$(0))
    REPOMINS$ = "10": REPOMIN = 10
    '
    Call GRACONTROL("CANECOM", "LOST-TER", LOST$(3))
    Call GRACONTROL("CANECOM", "LOST-MON", LOST$(2))
    Call GRACONTROL("CANECOM", "LOST-MEC", LOST$(1))
    Call GRACONTROL("CANECOM", "LOST-COM", LOST$(0))
    Call GRACONTROL("CANECOM", "REPOMIN", REPOMINS$)
    '
    FILPECLI$ = "NO": If Check1.Value = 1 Then FILPECLI$ = "SI"
    Call GRACONTROL("CANECOM", "FILPECLI", FILPECLI$)
    '
    MODO% = 0
    If Option1.Value = True Then MODO% = 1
    If Option2.Value = True Then MODO% = 2
    Call GRACONTROL("CANECOM", "MODOCALC", TRIM$(Str$(MODO%)))
    '
    Call GRACONTROL$("PROMATER", "DEPOSTOK", TRIM$(Text7))
    '
    PRISEMIN% = 53: If Option7.Value = True Then PRISEMIN% = 1
    Call GRACONTROL("", "PRISEMIN", TRIM$(Str$(PRISEMIN%)))
    '
End Sub

Private Sub Command3_Click()
    '
    Command3.Enabled = False
    If CLACONTRA% = 1 Then
        FF% = FECHANUM("17/10/13")
        FIN& = CantRegistros("ORDEFABR", "FECHGEN<'" & FETEXCOR1$(FF%) & "' AND STATUORF < 3")
        SQLX$ = "SELECT * FROM ORDEFABR WHERE FECHGEN<'" & FETEXCOR1$(FF%) & "' AND STATUORF < 3"
        Dim RST As ADODB.Recordset
        Set RST = New ADODB.Recordset
        RST.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic
        '
        With RST
            JJX& = 0
            Do While Not .EOF
                Call TRACE(24, JJX&, FIN&)
                FE% = CVINT(!FechEnSol)
                FL% = CVINT(!FechLan)
                FE1% = DIASPOST(FE%, 6)
                FL1% = DIASPOST(FL%, 30)
                '
                If FE1% < HOY(HO$) Then FE1% = FE%
                If FL1% < HOY(HO$) Then FL1% = FL%
                '
                !FechEnSol = CVDAT(FE1%)
                !FechLan = CVDAT(FL1%)
                .Update
                .MoveNext
                JJX& = JJX& + 1
            Loop
        End With
        '
        RST.Close
        Set RST = Nothing
    End If
    '
    Call COMUNI("Proceso Finalizado")
    '
    Command3.Enabled = True
    '
End Sub

Private Sub Command5_Click()
    '
    Command5.Enabled = False
    '
    If DERACCE%("ANUPROEN", "Anulación de Programas de Entregas") = 2 Then
    '
        Call SELECHO("SELFECHA/Anula Programa Anterior a...")
        FELIM$ = TRIM$(VALACT1$("SELFECHA"))
        If FELIM$ <> "" Then
          FECORTE = FECHANUM(FELIM$)
        Else
          GoTo 99
        End If
        
        Screen.MousePointer = 11
        APROGRA$ = "PROGENT"
        If MODOPLAN% = 1 Then APROGRA$ = "PROGFAB"   'MODOPLAN% variable global, se inicia en Form_Load()del formulario PROMAT07
        '                                            '0=programa entregas / 1=plan maestro de fabricacion
        FIN& = ULTREG&(APROGRA$)
        JJ& = 0
        For i& = 1 To FIN&
           Call TRACE(20, i&, FIN&)
           TXT$ = REGLEIDO$(APROGRA$, i&)
           FENTRE% = CVI(Mid$(TXT$, 3, 2))
           If FENTRE% >= FECORTE Then
             JJ& = JJ& + 1
             Call GRAREG(APROGRA$, TXT$, JJ&)
           End If
        Next i&
        Call SETULTREG(APROGRA$, JJ&)
        '
        Call PROMAT07.CARPROGENT
        Call PROMAT07.MUEPROGENT
        Call PROMAT07.GRAPROGENT
        Screen.MousePointer = 1
        '
        Call COMUNI("Proceso de Depuración Terminado")
    '
    End If
    '
99  Command5.Enabled = True
    '
End Sub

Private Sub Command9_Click()
    Call SELECHO("TADEPOSI")
    VDEVU$ = VALACT1$("TADEPOSI")
    If VDEVU$ <> "" Then
      Text7.TEXT = VDEVU$
    End If
End Sub

Private Sub Form_Activate()
    '
    Dim LOSTN(4), LOST$(4)
    '
    DoEvents
    CT$ = UCase$(Control$("CANECOM", "PERIODO"))
    SEMANAL% = 0: Option5.Value = True
    If InStr(CT$, "SEMAN") > 0 Then
        SEMANAL% = 1
        Option4.Value = True
    End If
    '
    CT$ = UCase$(Control$("CANECOM", "PLADISTA"))
    PLADISTA% = Val(CT$)
    If PLADISTA% < 0 Then PLADISTA% = 0
    If PLADISTA% > 6 Then PLADISTA% = 6
    Text1.TEXT = TRIM$(Str$(PLADISTA%))
    '
    CT$ = UCase$(Control$("CANECOM", "LOST-TER"))
    LOSTN(3) = Val(CT$): If LOSTN(3) < 0 Then LOSTN(3) = 0
    LOST$(3) = Mid$(Str$(LOSTN(3)), 2)
    Text2.TEXT = LOST$(3)
    '
    CT$ = UCase$(Control$("CANECOM", "LOST-MON"))
    LOSTN(2) = Val(CT$): If LOSTN(2) < 0 Then LOSTN(2) = 0
    LOST$(2) = Mid$(Str$(LOSTN(2)), 2)
    Text3.TEXT = LOST$(2)
    '
    CT$ = UCase$(Control$("CANECOM", "LOST-MEC"))
    LOSTN(1) = Val(CT$): If LOSTN(1) < 0 Then LOSTN(1) = 0
    LOST$(1) = Mid$(Str$(LOSTN(1)), 2)
    Text4.TEXT = LOST$(1)
    '
    CT$ = UCase$(Control$("CANECOM", "LOST-COM"))
    LOSTN(0) = Val(CT$): If LOSTN(0) < 0 Then LOSTN(0) = 0
    LOST$(0) = Mid$(Str$(LOSTN(0)), 2)
    Text5.TEXT = LOST$(0)
    '
    CT$ = UCase$(Control$("CANECOM", "REPOMIN"))
    REPOMIN = Val(CT$): If REPOMIN < 10 Then REPOMIN = 10
    REPOMINS$ = Mid$(Str$(REPOMIN), 2)
    '
    If UCase$(Control$("CANECOM", "FILPECLI")) = "SI" Then Check1.Value = 1
    '
    MODO% = Val(Control$("CANECOM", "MODOCALC"))
    If MODO% < 1 Or MODO% > 2 Then MODO% = 0
    Option3.Value = True
    If MODO% = 1 Then Option1.Value = True
    If MODO% = 2 Then Option2.Value = True
    '
    Text7 = TRIM$(Control$("PROMATER", "DEPOSTOK"))
    Label1 = ECOARCH$("TADEPOSI", Chr$(Val(Text7.TEXT)))
    '
    PRISEMIN% = 53
    If Val(Control$("", "PRISEMIN")) = 1 Then PRISEMIN% = 1
    If PRISEMIN% = 1 Then Option7.Value = True
    '
    If InStr(1, EMPREAC$, "FERVI") > 0 Then
        Command3.Visible = True
    End If
    '
    Refresh
    '
End Sub

Private Sub Form_Load()
   Call APLICACIONSKINS(Me, Picture1)
End Sub

Private Sub Text7_Change()
    If Val(Text7) < 1 Then
        Label1 = "Todos los Depósitos"
      Else
        Label1 = ECOARCH$("TADEPOSI", Chr$(Val(Text7)))
    End If
End Sub

Private Sub Text7_DblClick()
    Call SELECHO("TADEPOSI")
    VDEVU$ = VALACT1$("TADEPOSI")
    If VDEVU$ <> "" Then
      Text7.TEXT = VDEVU$
    End If
End Sub
