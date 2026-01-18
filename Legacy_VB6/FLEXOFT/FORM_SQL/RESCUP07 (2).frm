VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form RESCUP07 
   BackColor       =   &H00C0E0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resúmenes de Cuenta Corriente - Proveedores"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9300
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   9300
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   25
      Top             =   1320
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0C0&
      Caption         =   "Composición de Deuda"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   105
      TabIndex        =   23
      Top             =   2310
      Width           =   4005
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0E0C0&
         Caption         =   "Reconstrucción Automática"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   630
         TabIndex        =   24
         Top             =   400
         Width           =   2955
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00800000&
      Height          =   3480
      Left            =   8400
      ScaleHeight     =   3420
      ScaleWidth      =   915
      TabIndex        =   18
      Top             =   0
      Width           =   975
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   21
         Top             =   1260
         Width           =   645
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   22
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "O.K."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "RESCUP07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Emite Reporte Solicitado"
         Top             =   1800
         Width           =   650
      End
      Begin VB.CommandButton command1 
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
         Height          =   640
         Left            =   120
         Picture         =   "RESCUP07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Abandona la Aplicación"
         Top             =   315
         Width           =   650
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -315
         Picture         =   "RESCUP07.frx":0454
         Top             =   2630
         Width           =   1515
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0E0C0&
      Caption         =   "Rango de Fechas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   4320
      TabIndex        =   11
      Top             =   210
      Width           =   4005
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
         Height          =   360
         Left            =   2520
         TabIndex        =   16
         Top             =   630
         Width           =   1065
      End
      Begin VB.CommandButton Command10 
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
         Height          =   360
         Left            =   3570
         TabIndex        =   15
         Top             =   630
         Width           =   175
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   735
         TabIndex        =   13
         Top             =   630
         Width           =   1065
      End
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
         Height          =   360
         Left            =   1785
         TabIndex        =   12
         Top             =   630
         Width           =   175
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1995
         TabIndex        =   17
         Top             =   380
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Desde:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -1365
         TabIndex        =   14
         Top             =   380
         Width           =   2325
      End
   End
   Begin VB.ListBox List1 
      Height          =   540
      Left            =   8400
      Sorted          =   -1  'True
      TabIndex        =   10
      Top             =   2100
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0E0C0&
      Caption         =   "Rango de Cuentas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      Left            =   105
      TabIndex        =   1
      Top             =   210
      Width           =   4005
      Begin VB.TextBox Text10 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2625
         TabIndex        =   7
         Top             =   315
         Width           =   855
      End
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
         Height          =   360
         Left            =   3465
         TabIndex        =   6
         Top             =   315
         Width           =   175
      End
      Begin VB.TextBox Text9 
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
         Left            =   315
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   " "
         Top             =   735
         Width           =   3330
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2625
         TabIndex        =   4
         Text            =   " "
         Top             =   1155
         Width           =   855
      End
      Begin VB.CommandButton Command8 
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
         Height          =   360
         Left            =   3465
         TabIndex        =   3
         Top             =   1155
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
         Left            =   315
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1575
         Width           =   3330
      End
      Begin VB.Label Label6 
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
         Left            =   210
         TabIndex        =   9
         Top             =   420
         Width           =   2325
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta :"
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
         Left            =   210
         TabIndex        =   8
         Top             =   1260
         Width           =   2325
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0C0&
      Caption         =   "Selección y Ordenamiento"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1590
      Left            =   4320
      TabIndex        =   0
      Top             =   1540
      Width           =   4005
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0C0&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   360
         TabIndex        =   30
         Top             =   1000
         Width           =   3375
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0E0C0&
            Caption         =   "Numérico"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   120
            TabIndex        =   32
            Top             =   150
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0E0C0&
            Caption         =   "Alfabético"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   2
            Left            =   2000
            TabIndex        =   31
            Top             =   150
            Width           =   1380
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0E0C0&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   360
         TabIndex        =   27
         Top             =   240
         Width           =   3375
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0E0C0&
            Caption         =   "Cuentas de Gastos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Index           =   3
            Left            =   2000
            TabIndex        =   29
            Top             =   120
            Width           =   1275
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0E0C0&
            Caption         =   "Cuentas Corrientes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Index           =   0
            Left            =   120
            TabIndex        =   28
            Top             =   120
            Width           =   1500
         End
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   3990
         Y1              =   945
         Y2              =   945
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   750
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "RESCUP07.frx":2376
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   1590
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3600
      OleObjectBlob   =   "RESCUP07.frx":23F0
      Top             =   1200
   End
End
Attribute VB_Name = "RESCUP07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub
'

Private Sub Command10_Click()
    Call SELECHO("SELFECHA")
    If FECHANUM(VALACT1$("SELFECHA")) > 0 Then
        Text5.TEXT = VALACT1$("SELFECHA")
    End If
End Sub

Private Sub Command2_Click()
    CONTROLGRABA% = 0
    If Check1.Value = 1 Then CONTROLGRABA% = 1
    Hide
    Call FRESHALL
    Call GERESCUEPRO
    CONTROLGRABA% = 0
End Sub
'

Private Sub Command6_Click()
    Call SELECHO("SELFECHA")
    If FECHANUM(VALACT1$("SELFECHA")) > 0 Then
        Text6.TEXT = VALACT1$("SELFECHA")
    End If
End Sub

Private Sub Command7_Click()
   Call HELPONLINE("COGECOM")
End Sub

Private Sub Command8_Click()
    Call SELECHO("PROVED1")
    NCMAX% = XVALO(VALACT1$("PROVED1"))
    If NCMAX% > 0 Then
        Text8.TEXT = TRIM$(Str$(NCMAX%))
        Text7.TEXT = RASOCLI$((-1) * NCMAX%)
    End If
End Sub

Private Sub Command9_Click()
    Call SELECHO("PROVED1")
    NCMIN% = XVALO(VALACT1$("PROVED1"))
    If NCMIN% > 0 Then
        Text10.TEXT = TRIM$(Str$(NCMIN%))
        Text9.TEXT = RASOCLI$((-1) * NCMIN%)
    End If
End Sub

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    Call APLICACIONSKINS(Me, Picture1)
    '
End Sub

Private Sub Option1_Click(Index As Integer)
    If Index = 0 Then
      Option1(0) = False
      Option1(1) = True
    End If
End Sub

Private Sub Text10_Change()
    If Abs(XVALO(Text10)) < 32767 Then
       NCMIN% = XVALO(Text10.TEXT)
       Text9.TEXT = RASOCLI$((-1) * NCMIN%)
    End If
End Sub

Private Sub Text2_CLICK()
    Text1.SetFocus
End Sub


Private Sub Text5_DblClick()
    Call SELECHO("SELFECHA")
    If FECHANUM(VALACT1$("SELFECHA")) > 0 Then
      Text5.TEXT = VALACT1$("SELFECHA")
    End If
End Sub

Private Sub Text8_Change()
    If Abs(XVALO(Text8)) < 32767 Then
       NCMAX% = XVALO(Text8.TEXT)
       Text7.TEXT = RASOCLI$((-1) * NCMAX%)
    End If
End Sub

Private Sub Text10_DblClick()
    Call SELECHO("PROVED1")
    NCMIN% = XVALO(VALACT1$("PROVED1"))
    If NCMIN% > 0 Then
      Text10.TEXT = TRIM$(Str$(NCMIN%))
      Text9.TEXT = RASOCLI$((-1) * NCMIN%)
    End If
End Sub

Private Sub Text6_DblClick()
    Call SELECHO("SELFECHA")
    If FECHANUM(VALACT1$("SELFECHA")) > 0 Then
      Text6.TEXT = VALACT1$("SELFECHA")
    End If
End Sub

Private Sub Text7_CLICK()
    Text8.SetFocus
End Sub

Private Sub Text8_DblClick()
    Call SELECHO("PROVED1")
    NCMAX% = XVALO(VALACT1$("PROVED1"))
    If NCMAX% > 0 Then
      Text8.TEXT = TRIM$(Str$(NCMAX%))
      Text7.TEXT = RASOCLI$((-1) * NCMAX%)
    End If
End Sub

Private Sub Text9_CLICK()
    Text10.SetFocus
End Sub

Sub GERESCUEPRO()
    '
    If Abs(XVALO(Text10)) >= 32767 Or Abs(XVALO(Text8)) >= 32767 Then
       Call MENSERR(24, "Rango de Cuentas no Válido")
       Screen.MousePointer = 1
       Exit Sub
    End If
    '
    Screen.MousePointer = 11
    '
    DES% = FECHANUM(Text6.TEXT)
    HAS% = FECHANUM(Text5.TEXT)
    NCMIN% = XVALO(Text10.TEXT)
    NCMAX% = XVALO(Text8.TEXT)
    '
    MODORDE% = 1
    If Option1(2) = True Then MODORDE% = 2
    '
    FIN& = ULTREG&("PROVED1")
    UJ& = 0
    List1.Clear
    Call BLANARCH("RESCUPRO")
    '
    FIN& = ULTREG&("PROVED1")
    For U& = 1 To FIN&
      '
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("PROVED1", U&)
      NC% = CVI(Left$(X$, 2))
      '
      If NC% >= NCMIN% Then
        If NC% <= NCMAX% Then
          '
          TLISTA$ = Space$(36)
          If MODORDE% = 1 Then
               Call REPLA(TLISTA$, CSTRING$(MKI$(NC%), 1, 6, 0, 2), 1, 6)
             ElseIf MODORDE% = 2 Then
               Call REPLA(TLISTA$, RASOCLI$((-1) * NC%), 1, 30)
          End If
          Call REPLA(TLISTA$, CSTRING$(MKI$(NC%), 1, 6, 0, 2), 31, 6)
          List1.AddItem TLISTA$
          '
        End If
      End If
      '
    Next U&
    '
    FIN& = List1.ListCount
    '
    For U& = 1 To FIN&
      '
      Call TRACE(20, U&, FIN&)
      NC% = XVALO(Mid$(List1.List(U& - 1), 31, 6))
      RECUEP07.Text1 = TRIM$(Str$(NC%))
      If CONTROLGRABA% = 1 Then
        Call RECUEP07.Command4_Click
      End If
      'Call CARESCUEPRO
      Call FRESHALL
      '
      REBLA$ = REGBLAN$("RESCUPRO")
      CLIE$ = TRIM$(RECUEP07.Text2.TEXT) + " (" + TRIM$(Str$(NC%)) + ")"
      DOMIE$ = TRIM$(RECUEP07.Text3.TEXT)
      LOCAE$ = TRIM$(RECUEP07.Text5.TEXT)
      If TRIM$(RECUEP07.Text4.TEXT) <> "" Then
          LOCAE$ = TRIM$(RECUEP07.Text4.TEXT) + " - " + LOCAE$
      End If
      TELEE$ = "Telefax: " + TRIM$(RECUEP07.Text7.TEXT)
      '
      Call REPLA(REBLA$, MKI$(NC%), 1, 2)
      Call REPLA(REBLA$, "Sres.", 97, 6)
      Call REPLA(REBLA$, CLIE$, 103, 48)
      Call REPLA(REBLA$, DOMIE$, 151, 48)
      Call REPLA(REBLA$, LOCAE$, 199, 32)
      Call REPLA(REBLA$, TELEE$, 231, 24)
      '
      YATRANSPO% = 0: ITRANSPO# = 0
      For V& = 1 To RECUEP07.List1.ListCount
        RECUEP07.List1.ListIndex = V& - 1
        REGI$ = RECUEP07.List1.TEXT
        FECHA$ = Left$(REGI$, 8)
        FF% = FECHANUM(FECHA$)
        If FF% < DES% Then
            ITRANSPO# = XVALO(Mid$(REGI$, 64, 13))
        End If
        If FF% >= DES% Then
          If FF% <= HAS% Then
            If YATRANSPO% = 0 Then
              REGJ$ = REBLA$
              FEX = DES%
              Call REPLA(REGJ$, FECHATEX$(FEX), 3, 8)
              Call REPLA(REGJ$, "Saldo Anterior", 11, 32)
              Call REPLA(REGJ$, MKD$(ITRANSPO#), 63, 8)
              Call REGAPP("RESCUPRO", REGJ$)
              Call REGAPP("RESCUPRO", REBLA$)
              YATRANSPO% = 1
            End If
            '
            REGJ$ = REBLA$
            DOCU$ = Mid$(REGI$, 11, 11)
            NUDO$ = Mid$(REGI$, 22, 18)
            debe# = XVALO(Mid$(REGI$, 40, 11))
            HABER# = XVALO(Mid$(REGI$, 53, 11))
            SALDO# = XVALO(Mid$(REGI$, 64, 13))
            '
            Call REPLA(REGJ$, FECHA$, 3, 8)
            Call REPLA(REGJ$, DOCU$, 11, 18)
            Call REPLA(REGJ$, NUDO$, 29, 18)
            Call REPLA(REGJ$, MKD$(debe#), 47, 8)
            Call REPLA(REGJ$, MKD$(HABER#), 55, 8)
            Call REPLA(REGJ$, MKD$(SALDO#), 63, 8)
            Call REGAPP("RESCUPRO", REGJ$)
            '
          End If
        End If
      Next V&
      '
      If YATRANSPO% = 1 Then
          Call REGAPP("RESCUPRO", String$(256, "="))
      End If
      '
    Next U&
    Screen.MousePointer = 1
    '
    Call CONECRUN("*RESCPRO", "Listado Resumen de Cuenta Proveedores")
    '
End Sub
'





