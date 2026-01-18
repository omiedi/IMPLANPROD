VERSION 5.00
Begin VB.Form FORSELME 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Avance de Facturación y Cobranza"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
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
      Height          =   800
      Left            =   4245
      Picture         =   "FORSELME.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Terminar - Salir de la Aplicación"
      Top             =   315
      Width           =   960
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Opción"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2385
      Left            =   210
      TabIndex        =   9
      Top             =   945
      Width           =   3900
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Notas de Crédito"
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
         Left            =   630
         TabIndex        =   14
         Top             =   1470
         Value           =   1  'Checked
         Width           =   3060
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Notas de Débito"
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
         Left            =   630
         TabIndex        =   13
         Top             =   1155
         Value           =   1  'Checked
         Width           =   2955
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Facturas (Contado y C.Corriente)"
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
         Left            =   630
         TabIndex        =   12
         Top             =   840
         Value           =   1  'Checked
         Width           =   3165
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Cobranza por Recibo"
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
         TabIndex        =   11
         Top             =   1995
         Width           =   2850
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Facturación"
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
         TabIndex        =   10
         Top             =   465
         Value           =   -1  'True
         Width           =   1380
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Salida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   210
      TabIndex        =   6
      Top             =   3465
      Width           =   3900
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Pantalla"
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
         Left            =   630
         TabIndex        =   8
         Top             =   315
         Value           =   -1  'True
         Width           =   1380
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Impresión"
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
         Left            =   2205
         TabIndex        =   7
         Top             =   315
         Width           =   1275
      End
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   4245
      ScaleHeight     =   75
      ScaleWidth      =   915
      TabIndex        =   4
      Top             =   2415
      Width           =   980
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
      Height          =   620
      Left            =   4245
      Picture         =   "FORSELME.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Genera Reporte Solicitado"
      Top             =   3540
      Width           =   960
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
      Left            =   3930
      TabIndex        =   1
      Top             =   315
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
      Left            =   1725
      TabIndex        =   0
      Text            =   " "
      Top             =   315
      Width           =   2220
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   3760
      Picture         =   "FORSELME.frx":0454
      Top             =   2810
      Width           =   2010
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
      TabIndex        =   2
      Top             =   420
      Width           =   2220
   End
End
Attribute VB_Name = "FORSELME"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BOTREC_Click()
    '
    MMXX0$ = TRIM$(MESAN.TEXT)
    MMXX1$ = MESANO$(MMXX0$, DES%, HAS%)
    If TRIM$(MMXX1$) = "" Then
        Call MENSERR(24, "Ingrese el Mes a Listar")
        Exit Sub
    End If
    '
    Call BLANARCH("AVAFACB/NOERR")
    REAVA& = 0
    Screen.MousePointer = 11
    '\\\OT-05-0524-WAR-25/07/05///
    If Option3 Then
      If Option2 Then
        If EXISTE%(LUDAT$ + "FACBONIF.DAT") > 0 Then
          AX1% = COMALTER%("Formato de Salida: \\Normal\Detalle Bonificaciones")
          If AX1% = 2 Then
            Call LIFABON
            Screen.MousePointer = 1
            Exit Sub
           ElseIf AX1% = 0 Then
            Screen.MousePointer = 1
            Exit Sub
          End If
        End If
      End If
    End If
    '\\\OT-05-0524-WAR-FIN///
    LTX$ = LETREXPO$
    If Option3.Value = True Then
      FIN& = ULTREG&("FACTUR")
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
                '
                TI% = Asc(Mid$(X$, 13, 1)): If TI% < 0 Or TI% > 7 Then TI% = 2 ' default factura
                If TI% = 2 And Check1(0).Value <> 1 Then GoTo 19
                If TI% = 3 And Check1(1).Value <> 1 Then GoTo 19
                If TI% = 6 And Check1(2).Value <> 1 Then GoTo 19
                '
                VA% = Asc(Mid$(X$, 14, 1)): If VA% < 1 Or VA% > 3 Then VA% = 1 ' default factura
                SIG% = 1: If TI% > 4 Then SIG% = (-1)
                NETO# = Abs((Int(100 * (CVD(Mid$(X$, 15, 8)) + CVD(Mid$(X$, 49, 8))) + 0.5)) / 100)
                NETO# = NETO# * SIG%
                NUSUCUVE% = CVI(Mid$(X$, 47, 2))
                If TI% > 0 And TI% < 7 And TI% <> 4 Then
                    NNX$ = FORMDOC$(NRO&, TI%, VA%, NUSUCUVE%)
                    ' por VDR que ponga letra 'E' a exportaciones
                    PVCLX1% = PIVACLI%(NC%)
                    If PVCLX1% = 5 Or PVCLX1% = 7 Then
                      Call REPLA(NNX$, PREFEXPO$, 4, 4)
                      Call REPLA(NNX$, LETREXPO$, 18, 1)
                    End If
                    '
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
19    Next IL&
      '
    ElseIf Option4.Value = True Then ' Avance de Cobranza
      '
      FIN& = ULTREG&("COBRAN")
      For IL& = 1 To FIN&
        Call TRACE(20, IL&, FIN&)
        X$ = REGLEIDO$("COBRAN", IL&)
        FF% = CVI(Mid$(X$, 7, 2))
        If FF% >= DES% Then
            If FF% <= HAS% Then
                NC% = CVI(Left$(X$, 2))
                CLIE$ = RASOCLI$(NC%)
                '
                NUSUCUVE% = CVI(Mid$(X$, 23, 2))
                NRO& = CVS(Mid$(X$, 3, 4))
                NETO# = Abs((Int(100 * (CVD(Mid$(X$, 15, 8))) + 0.5)) / 100)
                PFX$ = TRIM$(Str$(NUSUCUVE%)): While Len(PFX$) < 2: PFX$ = "0" + PFX$: Wend
                NNX$ = TRIM$(Str$(NRO&)): While Len(NNX$) < 6: NNX$ = "0" + NNX$: Wend
                NNX$ = "RB-" + PFX$ + "-" + NNX$
                '
                Y$ = REGBLAN$("AVAFACB")
                Call REPLA(Y$, MKI$(FF%), 1, 2)
                Call REPLA(Y$, CLIE$, 3, 28)
                Call REPLA(Y$, NNX$, 31, 18)
                Call REPLA(Y$, MKD$(NETO#), 49, 8)
                REAVA& = REAVA& + 1
                Call GRAREG("AVAFACB", Y$, REAVA&)
            End If
        End If
      Next IL&
      '
    End If
    '
    Call SETULTREG("AVAFACB", REAVA&)
    Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    '
    Call FILESORT("AVAFACB", "", 1, 1, "ASC")
    '
    If Option1.Value = True Then
        '\\\OT-05-0524-WAR-02/08/05///
        If EXISTE%(LUDAT$ + "FACBONIF.DAT") > 0 Then
          AX1% = COMALTER%("Formato de Salida: \\Normal\Detalle Bonificaciones")
          If AX1% = 2 Then
            Call COFABONI
            Call FINAL("*COFABON")
            Screen.MousePointer = 1
            Exit Sub
           ElseIf AX1% = 0 Then
            Screen.MousePointer = 1
            Exit Sub
          End If
        End If
         '\\\OT-05-0524-WAR-FIN///
        Call FINAL("*COAVAFB")
      ElseIf Option2.Value = True Then
        Call FINAL("*LIAVAFB")
    End If
    '
    Hide
    '
End Sub

Private Sub Command1_Click()
   Call CIERRARCH("*.*")
   Hide
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

Private Sub Option3_Click()
   For KI% = 0 To 2
     Check1(KI%).Enabled = True
     Check1(KI%).Value = 1
   Next KI%
End Sub

Private Sub Option4_Click()
   For KI% = 0 To 2
     Check1(KI%).Value = 0
     Check1(KI%).Enabled = False
   Next KI%
End Sub
 '\\\OT-05-0524-WAR-25/07/05///
Sub LIFABON()
  '
   MXX0$ = TRIM$(MESAN.TEXT)
   MMXX1$ = MESANO$(MXX0$, DES%, HAS%)
   '
   Call COPYSTRU("FACBONIF", "LIFABONI")
   REAVA& = 0
   Screen.MousePointer = 11
   LTX$ = LETREXPO$
   '
   FIN& = ULTREG&("FACBONIF")
   For IL& = 1 To FIN&
     Call TRACE(20, IL&, FIN&)
     X$ = REGLEIDO$("FACBONIF", IL&)
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
         '
         TI% = Asc(Mid$(X$, 9, 1)): If TI% < 0 Or TI% > 7 Then TI% = 2 ' default factura
         If TI% = 2 And Check1(0).Value <> 1 Then GoTo 19
         If TI% = 3 And Check1(1).Value <> 1 Then GoTo 19
         If TI% = 6 And Check1(2).Value <> 1 Then GoTo 19
         '
         If TI% > 0 And TI% < 7 And TI% <> 4 Then
           NNX$ = Mid$(X$, 11, 18)
           ' por VDR que ponga letra 'E' a exportaciones
           PVCLX1% = PIVACLI%(NC%)
           If PVCLX1% = 5 Or PVCLX1% = 7 Then
             Call REPLA(NNX$, PREFEXPO$, 4, 4)
             Call REPLA(NNX$, LETREXPO$, 18, 1)
             Call REPLA(X$, NNX$, 11, 18)
           End If
           '
           REAVA& = REAVA& + 1
           Call GRAREG("LIFABONI", X$, REAVA&)
         End If
       End If
     End If
19 Next IL&
    '
  Call SETULTREG("LIFABONI", REAVA&)
  Call CIERRARCH("*.*")
  Screen.MousePointer = 1
  '
  Call FILESORT("LIFABONI", "LIFABONI", 6, 6, "ASC")
  Call FILESORT("LIFABONI", "LIFABONI", 3, 3, "ASC")
  '
  Call HEADERS("LIFABONI", "")
  Call HEADERS("LIFABONI", "Período: " + MXX0$)
  Call FINAL("*LIFABON")
  '
End Sub
 '\\\OT-05-0524-WAR-FIN///
 '\\\OT-05-0524-WAR-02/08/05///
Sub COFABONI()
  '
   MXX0$ = TRIM$(MESAN.TEXT)
   MMXX1$ = MESANO$(MXX0$, DES%, HAS%)
   '
   Call COPYSTRU("FACBONIF", "LIFABONI")
   REAVA& = 0
   Screen.MousePointer = 11
   LTX$ = LETREXPO$
   '
   FIN& = ULTREG&("FACBONIF")
   For IL& = 1 To FIN&
     Call TRACE(20, IL&, FIN&)
     X$ = REGLEIDO$("FACBONIF", IL&)
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
         '
         TI% = Asc(Mid$(X$, 9, 1)): If TI% < 0 Or TI% > 7 Then TI% = 2 ' default factura
         If TI% = 2 And Check1(0).Value <> 1 Then GoTo 19
         If TI% = 3 And Check1(1).Value <> 1 Then GoTo 19
         If TI% = 6 And Check1(2).Value <> 1 Then GoTo 19
         '
         If TI% > 0 And TI% < 7 And TI% <> 4 Then
           NNX$ = Mid$(X$, 11, 18)
           ' por VDR que ponga letra 'E' a exportaciones
           PVCLX1% = PIVACLI%(NC%)
           If PVCLX1% = 5 Or PVCLX1% = 7 Then
             Call REPLA(NNX$, PREFEXPO$, 4, 4)
             Call REPLA(NNX$, LETREXPO$, 18, 1)
             Call REPLA(X$, NNX$, 11, 18)
           End If
           '
           REAVA& = REAVA& + 1
           Call GRAREG("LIFABONI", X$, REAVA&)
         End If
       End If
     End If
19 Next IL&
    '
  Call SETULTREG("LIFABONI", REAVA&)
  Call CIERRARCH("*.*")
  Screen.MousePointer = 1
  '
  Call FILESORT("LIFABONI", "LIFABONI", 6, 6, "ASC")
  Call FILESORT("LIFABONI", "LIFABONI", 3, 3, "ASC")
  '
End Sub
 '\\\OT-05-0524-WAR-FIN///
