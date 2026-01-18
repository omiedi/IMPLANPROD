VERSION 5.00
Begin VB.Form FORSELMES 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleccion Periodo Mensual"
   ClientHeight    =   1410
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   315
      ScaleHeight     =   75
      ScaleWidth      =   3600
      TabIndex        =   4
      Top             =   1060
      Width           =   3660
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.CommandButton BOTREC 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   760
      Left            =   4305
      Picture         =   "FORSELMES.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Genera Reporte Solicitado"
      Top             =   420
      Width           =   750
   End
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
      Height          =   360
      Left            =   3780
      TabIndex        =   1
      Top             =   420
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
      Height          =   360
      Left            =   1575
      TabIndex        =   0
      Text            =   " "
      Top             =   420
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
      Left            =   -735
      TabIndex        =   2
      Top             =   525
      Width           =   2220
   End
End
Attribute VB_Name = "FORSELMES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BOTREC_Click()
    MMXX0$ = TRIM$(MESAN.TEXT)
    MMXX1$ = MESANO$(MMXX0$, DES%, HAS%)
    If TRIM$(MMXX1$) = "" Then
        Exit Sub
    End If
    '
    REAVA& = 0
    FIN& = ULTREG&("FACTUR")
    '
    Screen.MousePointer = 11
    For IL& = 1 To FIN&
        Call TRACE(20, IL&, FIN&)
        X$ = REGLEIDO$("FACTUR", IL&)
        FF% = CVI(Mid$(X$, 7, 2))
        If FF% >= DES% Then
            If FF% <= HAS% Then
                NC% = CVI(Left$(X$, 2))
                CLIE$ = ""
                If NC% > 0 Then
                      CLIE$ = RASOCLI$(NC%)
                   Else
                      REGCU& = Abs(NC%)
                      If REGCU& > 0 Then
                         If REGCU& <= ULTREG&("CUITS") Then
                            Z$ = REGLEIDO$("CUITS", REGCU&)
                            CLIE$ = Left$(Z$, 30)
                         End If
                      End If
                End If
                NRO& = CVS(Mid$(X$, 3, 4))
                TI% = Asc(Mid$(X$, 13, 1)): If TI% < 0 Or TI% > 7 Then TI% = 2 ' default factura
                VA% = Asc(Mid$(X$, 14, 1)): If VA% < 1 Or VA% > 3 Then VA% = 1 ' default factura
                SIG% = 1: If TI% > 4 Then SIG% = (-1)
                NETO# = Abs((Int(100 * (CVD(Mid$(X$, 15, 8)) + CVD(Mid$(X$, 57, 8))) + 0.5)) / 100)
                NETO# = NETO# * SIG%
                If TI% > 0 And TI% < 7 And TI% <> 4 Then
                    NNX$ = FORMDOC$(NRO&, TI%, VA%, 1)
                    Y$ = REGBLAN$("AVAFACB")
                    Call REPLA(Y$, MKI$(FF%), 1, 2)
                    Call REPLA(Y$, CLIE$, 3, 28)
                    Call REPLA(Y$, NNX$, 31, 18)
                    Call REPLA(Y$, MKD$(NETO#), 49, 8)
                    REAVA& = REAVA& + 1
                    Call GRAREG("AVAFACB", Y$, REAVA&)
                End If
            End If
        End If
    Next IL&
    '
    Call SETULTREG("AVAFACB", REAVA&)
    Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    '
    If COGEVENT.Option1.Value = True Then
         Call FINAL("*COAVAFB")
      ElseIf COGEVENT.Option2.Value = True Then
         Call FINAL("*SOAVAFB")
    End If
    '
    Hide
    '
End Sub

Private Sub Command3_Click()
    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) = "" Then
        Exit Sub
    End If
    MESAN.TEXT = MESASEL$
End Sub

Private Sub Form_Load()
   '

   Call TRANSLABELS(Name)
   Call APLICACIONSKINS(Me)

   '
End Sub

