VERSION 5.00
Begin VB.Form MABRUNOV 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Analisis de Rentabilidad"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   7305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   7305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6195
      Picture         =   "MABRUNOV.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Terminar - Salir de Consulta de Costos"
      Top             =   345
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Emision de Listados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   105
      TabIndex        =   12
      Top             =   2415
      Width           =   5895
      Begin VB.CommandButton Command6 
         Height          =   645
         Left            =   2520
         Picture         =   "MABRUNOV.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Ranking por Cliente y Margen Total"
         Top             =   420
         Width           =   645
      End
      Begin VB.CommandButton Command4 
         Height          =   645
         Left            =   1785
         Picture         =   "MABRUNOV.frx":058C
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Ranking por Producto y Margen Total"
         Top             =   420
         Width           =   645
      End
      Begin VB.CommandButton Command2 
         Height          =   645
         Left            =   1050
         Picture         =   "MABRUNOV.frx":0896
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Ranking por Producto y Margen Procentual"
         Top             =   420
         Width           =   645
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Avance de Proceso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   3360
         TabIndex        =   14
         Top             =   420
         Width           =   2325
         Begin VB.PictureBox MARCOBARRA 
            BackColor       =   &H000000FF&
            Height          =   135
            Left            =   315
            ScaleHeight     =   75
            ScaleWidth      =   1605
            TabIndex        =   15
            Top             =   315
            Width           =   1665
            Begin VB.Frame BARRATRAZA 
               BackColor       =   &H00FF0000&
               BorderStyle     =   0  'None
               Height          =   500
               Left            =   0
               TabIndex        =   16
               Top             =   0
               Width           =   12000
            End
         End
      End
      Begin VB.CommandButton Command1 
         Height          =   645
         Left            =   315
         Picture         =   "MABRUNOV.frx":0BA0
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Listado Secuencial por Número de Comprobante"
         Top             =   420
         Width           =   645
      End
   End
   Begin VB.CommandButton BOTREC 
      Caption         =   "Re-Calc"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   4830
      Picture         =   "MABRUNOV.frx":0EAA
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Genera Reporte Solicitado"
      Top             =   840
      Width           =   855
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3885
      Left            =   -2940
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   3780
      Visible         =   0   'False
      Width           =   11775
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3885
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   3675
      Visible         =   0   'False
      Width           =   11775
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Período Activo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      Left            =   105
      TabIndex        =   0
      Top             =   210
      Width           =   5895
      Begin VB.CommandButton Command3 
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
         Height          =   400
         Left            =   3915
         TabIndex        =   2
         Top             =   620
         Width           =   175
      End
      Begin VB.TextBox MESAN 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1680
         TabIndex        =   1
         Text            =   " "
         Top             =   630
         Width           =   2250
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "hs"
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
         Index           =   2
         Left            =   4200
         TabIndex        =   11
         Top             =   1425
         Width           =   540
      End
      Begin VB.Label Label3 
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
         Height          =   300
         Left            =   3360
         TabIndex        =   10
         Top             =   1365
         Width           =   750
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "a las"
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
         Index           =   1
         Left            =   2415
         TabIndex        =   9
         Top             =   1425
         Width           =   855
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
         Height          =   300
         Left            =   1680
         TabIndex        =   8
         Top             =   1365
         Width           =   1065
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Generado el:"
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
         Index           =   0
         Left            =   -630
         TabIndex        =   7
         Top             =   1430
         Width           =   2220
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Mes y año:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   10
         Left            =   -630
         TabIndex        =   3
         Top             =   735
         Width           =   2220
      End
   End
   Begin VB.Image Image2 
      Height          =   555
      Left            =   5985
      Picture         =   "MABRUNOV.frx":11B4
      Stretch         =   -1  'True
      Top             =   3110
      Width           =   5355
   End
End
Attribute VB_Name = "MABRUNOV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BOTREC_Click()
    Call CALMANOVO
End Sub

Private Sub Command1_Click()
   Call FINAL("*LIAMAVE")
End Sub

Private Sub Command2_Click()
   Call RAMAVEN(1)
End Sub

Private Sub Command3_Click()
    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) = "" Then
        Exit Sub
    End If
    MESAN.Text = MESASEL$
End Sub

Private Sub Command4_Click()
   Call RAMAVEN(2)
End Sub

Private Sub Command5_Click()
   Call CIERRARCH("*.*")
   Unload MABRUNOV
End Sub

Private Sub Command6_Click()
   Call RAMAVEN(3)
End Sub

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    Call DATARGEN("AMAVEN", FEGE%, HOGE%, USGE%, TEGE%, CORRE$)
    '
    FEX = FEGE%
    Label1.Caption = FECHATEX$(FEX)
    Label3.Caption = HORATEX$(HOGE%)
    MESAN.Text = CORRE$
    '
    If TRIM$(MESAN.Text) <> "" Then
      If FECHANUM(Label1.Caption) > 0 Then
        Command1.Enabled = True
        Command4.Enabled = True
        Command2.Enabled = True
      End If
    End If
    '
End Sub

Private Sub MESAN_Change()
    Command1.Enabled = False
    Command4.Enabled = False
    Command2.Enabled = False
End Sub

Sub RAMAVEN(MODO%)
    '
    
    XX$ = Left$(REGLEIDO$("ECOMAST", 1), 16): CI1% = Abs(CODINT%(XX$))
    XX$ = Left$(REGLEIDO$("ECOMAST", ULTREG&("ECOMAST")), 16): CI2% = Abs(CODINT%(XX$))
    VDEF$ = MKI$(CI1%) + MKI$(CI2%) + Chr$(0)
    RACOD$ = VDEF$
    If MODO% < 3 Then
10    RACOD$ = OBJPLA$("DESHASCOD1", VDEF$)
      If RACOD$ = "" Then
        Exit Sub
      End If
    End If
    CI1% = CVI(Left$(RACOD$, 2)): cod1$ = CODEXT$(CI1%)
    CI2% = CVI(Mid$(RACOD$, 3, 2)): cod2$ = CODEXT$(CI2%)
    TMT% = Asc(Mid$(RACOD$, 5, 1))
    TIMAT$ = ECOARCH$("TATIMAT", Chr$(TMT%))
    If TIMAT$ = "" Then
        TMT% = 0
        Call REPLA(RACOD$, Chr$(0), 5, 1)
        TIMAT$ = "Sin Clasificación"
    End If
    '
    If MODO% < 3 Then
      If cod1$ = "" Or cod2$ = "" Or cod2$ < cod1$ Then
        VDEF$ = RACOD$
        GoTo 10
      End If
    End If
    '
    Screen.MousePointer = 11
    Dim PUNMAS$(32767)
    '
    SUFAC# = 0
    FIN& = ULTREG&("AMAVEN")
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = REGLEIDO$("AMAVEN", U&)
       CIXI% = CVI(Mid$(X$, 9, 2))
       SUFAC# = SUFAC# + Abs(CVD(Mid$(X$, 43, 8)))
       If MODO% = 3 Then CIXI% = CVI(Mid$(X$, 7, 2))
       If CIXI% > 0 Then
         If CIXI% <= 32767 Then
           PUNMAS$(CIXI%) = PUNMAS$(CIXI%) + MKS$(U&)
         End If
       End If
    Next U&
    '
    ALEA$ = "INVERT": INI& = 2: FIN& = NUINV% + 1: ASAL$ = "RAMAVEN"
    If MODO% = 3 Then
       ALEA$ = "DEUDOR1": INI& = 1: FIN& = ULTREG&("DEUDOR1"): ASAL$ = "DAMAVEN"
    End If
    REGQ& = 0
    For U& = 1 To FIN&
        Call TRACE(20, U&, FIN&)
        X$ = REGLEIDO$(ALEA$, U&)
        CODYY$ = Left$(X$, 16)
        If MODO% = 3 Then CODYY$ = AJUSTI$(cod1$, 16)
        If CODYY$ >= cod1$ Then
          If CODYY$ <= cod2$ Then
            If Mid$(CODYY$, 16, 1) <> Chr$(96) Then
              CIXI% = CVI(Mid$(X$, 17, 2))
              If MODO% = 3 Then CIXI% = CVI(Left$(X$, 2))
              If CIXI% > 0 Then
                If CIXI% <= 32767 Then
                  If TMT% = 0 Or TIMATE%(CIXI%) = TMT% Then
                    If Len(PUNMAS$(CIXI%)) > 0 Then
                      ITOT# = 0
                      IPRI# = 0
                      IMOD# = 0
                      ITER# = 0
                      For K% = 1 To Len(PUNMAS$(CIXI%)) Step 4
                        REG& = CVS(Mid$(PUNMAS$(CIXI%), K%, 4))
                        Z$ = REGLEIDO$("AMAVEN", REG&)
                        IPRI# = IPRI# + CVS(Mid$(Z$, 65, 4))
                        IMOD# = IMOD# + CVS(Mid$(Z$, 69, 4))
                        ITER# = ITER# + CVS(Mid$(Z$, 73, 4))
                        ITOT# = ITOT# + CVD(Mid$(Z$, 43, 8))
                      Next K%
                      '
                      IMAR# = ITOT# - IPRI# - IMOD# - ITER#
                      Y$ = String$(64, 0)
                      PORPRI = 0: If ITOT# > 0 Then PORPRI = 100 * IPRI# / ITOT#
                      PORMOD = 0: If ITOT# > 0 Then PORMOD = 100 * IMOD# / ITOT#
                      PORTER = 0: If ITOT# > 0 Then PORTER = 100 * ITER# / ITOT#
                      PORMAR = 0: If ITOT# > 0 Then PORMAR = 100 * IMAR# / ITOT#
                      Call REPLA(Y$, MKI$(CIXI%), 1, 2)
                      Call REPLA(Y$, MKD$(ITOT#), 3, 8)
                      Call REPLA(Y$, MKD$(IPRI#), 11, 8)
                      Call REPLA(Y$, MKS$(PORPRI), 19, 4)
                      Call REPLA(Y$, MKD$(IMOD#), 23, 8)
                      Call REPLA(Y$, MKS$(PORMOD), 31, 4)
                      Call REPLA(Y$, MKD$(ITER#), 35, 8)
                      Call REPLA(Y$, MKS$(PORTER), 43, 4)
                      Call REPLA(Y$, MKD$(IMAR#), 47, 8)
                      Call REPLA(Y$, MKS$(PORMAR), 55, 4)
                      '
                      If MODO% = 1 Then
                         ISOR = 100 - PORMAR
                        Else
                         ISOR = SUFAC# - IMAR#
                      End If
                      '
                      REGQ& = REGQ& + 1
                      Call GRAREG(ASAL$, MKS$(ISOR) + Y$, REGQ&)
                      '
                    End If
                  End If
                End If
              End If
            End If
          End If
        End If
    Next U&
    '
    Call SETULTREG(ASAL$, REGQ&)
    Call CIERRARCH(ASAL$)
    '
    Call CONECRUN("*SO" + Left$(ASAL$, 5), "Ranking por Margen de Ventas")
    '
End Sub
