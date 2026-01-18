VERSION 5.00
Begin VB.Form CIERRCTEL 
   BackColor       =   &H80000016&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cierre de Consultas Telefonicas"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10380
   Icon            =   "CIERRCTEL.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   10380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   9555
      TabIndex        =   13
      Top             =   3990
      Width           =   645
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   9555
      TabIndex        =   12
      Top             =   3675
      Width           =   645
   End
   Begin VB.Frame Frame4 
      Caption         =   "Procesar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1420
      Left            =   8025
      TabIndex        =   9
      Top             =   2860
      Width           =   1260
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
         Height          =   525
         Left            =   150
         Picture         =   "CIERRCTEL.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Procesa Cierre General"
         Top             =   735
         Width           =   960
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   150
         TabIndex        =   10
         Top             =   315
         Width           =   960
      End
   End
   Begin VB.CommandButton BOTEXIT 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   9450
      Picture         =   "CIERRCTEL.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Cancela los Datos Ingresados"
      Top             =   2940
      Width           =   750
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000013&
      Caption         =   "Consultas pendientes de cierre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2700
      Left            =   105
      TabIndex        =   6
      Top             =   0
      Width           =   10095
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2100
         Left            =   210
         Style           =   1  'Checkbox
         TabIndex        =   7
         Top             =   420
         Width           =   9650
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      Caption         =   "  Consulta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1440
      Left            =   105
      TabIndex        =   4
      Top             =   2835
      Width           =   7785
      Begin VB.TextBox Text22 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   210
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   315
         Width           =   7380
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000013&
      Caption         =   "  Diagnóstico e Instrucciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3430
      Left            =   105
      TabIndex        =   2
      Top             =   4335
      Width           =   10095
      Begin VB.TextBox Text23 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2955
         Left            =   210
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   315
         Width           =   9650
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H80000013&
      Caption         =   "Número"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   8400
      TabIndex        =   0
      Top             =   1890
      Visible         =   0   'False
      Width           =   1695
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   105
         TabIndex        =   1
         Top             =   315
         Width           =   1485
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   225
      Left            =   5250
      TabIndex        =   17
      Top             =   2730
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   225
      Left            =   3990
      TabIndex        =   16
      Top             =   2730
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "P:"
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
      Left            =   9240
      TabIndex        =   15
      Top             =   4060
      Width           =   300
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "T:"
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
      Left            =   9240
      TabIndex        =   14
      Top             =   3750
      Width           =   300
   End
End
Attribute VB_Name = "CIERRCTEL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NUCETE$
Dim ABUSQUE$

Private Sub Command10_Click()
   Call PRIORTRA(TRIM$(Label15.Caption))
End Sub

Private Sub Command13_Click()
   Call SELECHO("SELFECHA")
   VDEVU$ = VALACT1$("SELFECHA")
   If VDEVU$ <> "" Then
     Text26 = VDEVU$
   End If
End Sub

Private Sub Command12_Click()
    '
    FORSELEC.ABUSQUE = Label15
    Call SELECHO("CORTRAB1")
    NUCOTE$ = AJUSTI$(VALACT1$("CORTRAB1"), 12)
    If TRIM$(NUCOTE$) <> "" Then
       Label15.Caption = NUCOTE$
    End If
    '
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   Call CONECRUN("FLSETUP", "Configuración de Funcionamiento")
   Call FINAL("")
   Command14.Enabled = True
End Sub

Private Sub Command1_Click()
    '
    Call SELECHO("DEUDOR1")
    NCXI% = Val(TRIM$(VALACT1$("DEUDOR1")))
    If NCXI% > 0 Then
      NC% = NCXI%
      Text1.TEXT = TRIM$(Str$(NC%))
      Text2.TEXT = RASOCLI$(NC%)
      Text3.TEXT = TRIM$(DOMICLI$(NC%)) + " - (" + TRIM$(CPOSCLI$(NC%)) + ") " + LOCACLI$(NC%)
      Text4.TEXT = TEL1CLI$(NC%)
      Text5.TEXT = FAXCLI$(NC%)
      Text10 = ""
      Text11 = ""
      '
      For U& = ULTREG&("COTELEF1") To 1 Step -1
        XX$ = REGLEIDO$("COTELEF1", U&)
        If CVI(Mid$(XX$, 81, 2)) = NCXI% Then
          Text11 = TRIM$(Mid$(XX$, 83, 24))
          CCXX$ = TRIM$(Mid$(XX$, 179, 24))
          If TRIM$(Mid$(XX$, 203, 24)) <> "" Then CCXX$ = CCXX$ + CRLF$ + TRIM$(Mid$(XX$, 203, 24))
          If TRIM$(Mid$(XX$, 227, 24)) <> "" Then CCXX$ = CCXX$ + CRLF$ + TRIM$(Mid$(XX$, 227, 24))
          Text10 = CCXX$
          Exit Sub
        End If
      Next U&
    End If
    '
End Sub

Private Sub Command11_Click()
    Command11.Enabled = False
    Call CIERRARCH("*.*")
    ACONEC$ = "AMACLI02"
    If EXISTE%(ACONEC$ + ".EXE") < 1 Then
      ACONEC$ = "AMACLI01"
      If EXISTE%(ACONEC$ + ".EXE") < 1 Then
        ACONEC$ = "AMACLI"
      End If
    End If
    Call CONECRUN(ACONEC$, "Alta de Nueva Cuenta de Cliente")
    Call FRESHECHO("DEUDOR1")
    Command11.Enabled = True
End Sub

Private Sub Command15_Click()
   Call SELECHO("SELFECHA")
   VDEVU$ = VALACT1$("SELFECHA")
   If VDEVU$ <> "" Then
     Text27 = VDEVU$
   End If
End Sub

Private Sub Command16_Click()
   If USNBR% <> 101 Then Exit Sub
   Call SELECHO("SELFECHA")
   VDEVU$ = VALACT1$("SELFECHA")
   If VDEVU$ <> "" Then
     Text20 = VDEVU$
   End If
End Sub

Private Sub Command17_Click()
    '
    Command17.Enabled = False
    Screen.MousePointer = 11
    JJ& = 0
    Call COPYSTRU("CORTRAB1", "CORTRASE")
    '
    For U& = 1 To ULTREG("CORTRAB1")
      XX$ = REGLEIDO$("CORTRAB1", U&)
      If CVI(Mid$(XX$, 115, 2)) < 1 Then
        JJ& = JJ& + 1
        Call GRAREG("!CORTRASE", XX$, JJ&)
      End If
    Next U&
    Call SETULTREG("!CORTRASE", JJ&)
    Call FRESHECHO("!CORTRASE")
    Screen.MousePointer = 1
    '
    FORSELEC.ABUSQUE = Label15
    Call SELECHO("!CORTRASE")
    NUCOTE$ = AJUSTI$(VALACT1$("!CORTRASE"), 12)
    If TRIM$(NUCOTE$) <> "" Then
       Label15.Caption = NUCOTE$
    End If
    '
    Command17.Enabled = True
    '
End Sub

Private Sub Command18_Click()
   Command18.Enabled = False
   Call FINAL("*TAROTRA")
   Command18.Enabled = True
End Sub

Private Sub Command19_Click()
   Command19.Enabled = False
   Dim RENGLO$(32)
   REBLA$ = REGBLAN$("LIOTPEN")
   OPX% = ALTERNA%("Pendientes Por Cliente\Pendientes por Fecha")
   If OPX% = 1 Then
       Call COPYSTRU("CORTRAB1", "CORTRAB2")
       JJ& = 0
       For KU& = 1 To ULTREG&("CORTRAB1")
         XX$ = REGLEIDO$("CORTRAB1", KU&)
         If CVI(Mid$(XX$, 115, 2)) < 1 Then   ' NO ESTA CERRADA
           JJ& = JJ& + 1
           Call GRAREG("CORTRAB2", XX$, JJ&)
         End If
       Next KU&
       Call SETULTREG("CORTRAB2", JJ&)
       '
       Call FILESORT("CORTRAB2", "", 4, 5, "ASC")
       JJ& = 0
       NORI& = 0
       For KU& = 1 To ULTREG&("CORTRAB2")
         XX$ = REGLEIDO$("CORTRAB2", KU&)
         NORI& = NORI& + 1
         NC% = CVI(Mid$(XX$, 81, 2))
         '
         If NC% > 0 Then
           DECLI$ = TEL2CLI$(NC%)
           DETA$ = TRIM$(Mid$(XX$, 321, 256))
           FALTA% = CVI(Mid$(XX$, 107, 2))
           OBSX$ = TRIM$(Mid$(XX$, 193, 128))
           '
           For KKL& = 1 To 32
             RENGLO$(KKL&) = REBLA$
           Next KKL&
           '
           NORI& = NORI& + 1
           Call FRATEXTO(DETA$, 64)
           For LL% = 1 To CARETEX%
             TTYY$ = RENGLO$(LL%)
             Call REPLA(TTYY$, MKS$(NORI&), 1, 4)
             Call REPLA(TTYY$, DECLI$, 5, 12)
             Call REPLA(TTYY$, RETEX$(LL%), 17, 64)
             Call REPLA(TTYY$, MKI$(FALTA%), 81, 2)
             RENGLO$(LL%) = TTYY$
           Next LL%
           CAREMA% = CARETEX%
           '
           Call FRATEXTO(OBSX$, 40)
           For LL% = 1 To CARETEX%
             TTYY$ = RENGLO$(LL%)
             Call REPLA(TTYY$, MKS$(NORI&), 1, 4)
             Call REPLA(TTYY$, DECLI$, 5, 12)
             Call REPLA(TTYY$, MKI$(FALTA%), 81, 2)
             Call REPLA(TTYY$, RETEX$(LL%), 83, 40)
             RENGLO$(LL%) = TTYY$
           Next LL%
           If CARETEX% > CAREMA% Then CAREMA% = CARETEX%
           '
           For LL% = 1 To CAREMA%
             JJ& = JJ& + 1
             Call GRAREG("LIOTPEN", RENGLO$(LL%), JJ&)
           Next LL%
         End If
       Next KU&
       Call SETULTREG("LIOTPEN", JJ&)
       Call FILESORT("LIOTPEN", "", 2, 2, "ASC")
       Call CIERRARCH("LIOTPEN")
       Call FINAL("*LIOTPEN")
       '
   End If
   '
   Command19.Enabled = True
End Sub

Private Sub Command20_Click()
   Command20.Enabled = False
   ADJUNARCH.Show 1
   Command20.Enabled = True
End Sub

Function PROTARIN&()
   '
   PROTARIN& = 1
   URE& = ULTREG&("TARINTER")
   If URE& > 0 Then
     F$ = REGLEIDO$("TARINTER", URE&)
     PROTARIN& = 1 + CVS(Left$(F$, 4))
   End If
   '
End Function

'
Private Sub BOTEXIT_Click()
    Unload Me
End Sub

Private Sub BOTREC_Click()
   '
   Screen.MousePointer = 11
   For KU& = 1 To List1.ListCount
     If List1.Selected(KU& - 1) = True Then
        NUCOTE$ = Left$(List1.List(KU& - 1), 12)
        CICOTE% = 0: For UI% = 1 To Len(NUCOTE$)
                       CICOTE% = CICOTE% + UI% * Asc(Mid$(NUCOTE$, UI%, 1))
                     Next UI%
                     '
        RNC$ = RECFILT$("COTELEF1", 3, MKI$(CICOTE%))
        URN& = ULTREG&("COTELEF1")
        '
        For U& = 1 To Len(RNC$) Step 4
          RENOTA& = CVS(Mid$(RNC$, U&, 4))
          If RENOTA& > 0 Then
            If RENOTA& <= URN& Then
              TTXX$ = REGLEIDO$("COTELEF1", RENOTA&)
              If Left$(TTXX$, 12) = NUCOTE$ Then
                FEXI% = FECHANUM(Text1)
                Call REPLA(TTXX$, MKI$(FEXI%), 251, 2)
                Call GRAREG("COTELEF1", TTXX$, RENOTA&)
                GoTo 29
              End If
            End If
          End If
        Next U&
        '
      End If
29 Next KU&
   '
   JJ& = 0
   For U& = 1 To ULTREG("COTELEF1")
     TTXX$ = REGLEIDO$("COTELEF1", U&)
     If CVI(Mid$(TTXX$, 251, 2)) <= 8224 Then
       JJ& = JJ& + 1
       Call GRAREG("COTELEF3", TTXX$, JJ&)
     End If
   Next U&
   Call SETULTREG("COTELEF3", JJ&)
   Call CIERRARCH("COTELEF3")
   Call FRESHECHO("COTELEF3")
   '
   Call CIERRARCH("COTELEF1")
   '
   Screen.MousePointer = 1
   Call COMUNI("Cierre Procesado")
   Unload Me
   '
End Sub

Private Sub Form_Load()
    Text1.TEXT = HOS$
End Sub

Private Sub Label7_Change()
   Call CARDATCOTE
   Call CARDETCOTE
End Sub

Private Sub List1_Click()
   Label7 = Left$(List1.TEXT, 12)
   
End Sub

Private Sub List1_DblClick()
    CTACT$ = TRIM$(Left$(List1.TEXT, 12))
    If COMALTER%("¿ Está Seguro de Eliminar la C-Tel '" + CTACT$ + "' ?\\ No, Cancelar\Si, Eliminar") = 2 Then
       Call CIERRARCH("COTELEF1")
       Call BLOQARCH("COTELEF1")
       Screen.MousePointer = 11
       Call FILTERDELETE("COTELEF1", 1, CTACT$)
       '
       DES% = XVALO(Label2)
       HAS% = XVALO(Label3)
       JJ& = 0
       TOT& = 0: PEN& = 0
       CIERRCTEL.List1.Clear     ' EPB - NECESARIO PORQUE SI NO AL VOLVER A ENTRAR DUPLICA LA CARGA
       For U& = 1 To ULTREG&("COTELEF1")
         XX$ = REGLEIDO$("COTELEF1", U&)
         FEMI% = CVI(Mid$(XX$, 107, 2))
         FECI% = CVI(Mid$(XX$, 251, 2))
         If FEMI% >= DES% Then
           If FEMI% <= HAS% Then
             TOT& = TOT& + 1
             If FECI% <= 8224 Then          ' NO ESTÁ CERRADA
               PEN& = PEN& + 1
               NUCOTE$ = Mid$(XX$, 1, 12)
               FECLIRE$ = Mid$(XX$, 13, 66)
               CIERRCTEL.List1.AddItem (NUCOTE$ + " - " + FECLIRE$)
             End If
           End If
         End If
       Next U&
       '
       CIERRCTEL.Text2 = TOT&
       CIERRCTEL.Text3 = PEN&
       Screen.MousePointer = 1
    End If
    '
End Sub

Private Sub Text1_DblClick()
    '
    Call SELECHO("SELFECHA")
    Text1 = VALACT1$("SELFECHA")
    
End Sub

Sub CARDETCOTE()
   '
   NUCOTE$ = AJUSTI$(Label7.Caption, 12)
   CICOTE% = 0: For UI% = 1 To Len(NUCOTE$)
                  CICOTE% = CICOTE% + UI% * Asc(Mid$(NUCOTE$, UI%, 1))
                Next UI%
                '
   Screen.MousePointer = 11
   URN& = ULTREG&("COTELEF2")
   Text22.TEXT = "": Text23.TEXT = ""
   '
   Screen.MousePointer = 11
   RNC$ = RECFILT$("COTELEF2", 2, MKI$(CICOTE%))
   '
   For U& = 1 To Len(RNC$) Step 4
     RENOTA& = CVS(Mid$(RNC$, U&, 4))
     If RENOTA& > 0 Then
       If RENOTA& <= URN& Then
         TTXX$ = REGLEIDO$("COTELEF2", RENOTA&)
         If Left$(TTXX$, 12) = NUCOTE$ Then
           TIRE$ = Mid$(TTXX$, 15, 2)
           TTYY$ = Mid$(TTXX$, 17, 72)
           If TIRE$ = "CO" Then
                If Text22.TEXT <> "" Then
                  Text22.TEXT = Text22.TEXT + Chr$(13) + Chr$(10)
                End If
                Text22.TEXT = Text22.TEXT + RTrim$(Left$(TTYY$, 72))
              ElseIf TIRE$ = "DI" Then
                If Text23.TEXT <> "" Then
                  Text23.TEXT = Text23.TEXT + Chr$(13) + Chr$(10)
                End If
                Text23.TEXT = Text23.TEXT + RTrim$(Left$(TTYY$, 72))
           End If
         End If
       End If
     End If
   Next U&
   Screen.MousePointer = 1
   Text22.Refresh: Text23.Refresh
   '
End Sub

Sub CARDATCOTE()
   '
   NUCOTE$ = AJUSTI$(Label7.Caption, 12)
   CICOTE% = 0: For UI% = 1 To Len(NUCOTE$)
                  CICOTE% = CICOTE% + UI% * Asc(Mid$(NUCOTE$, UI%, 1))
                Next UI%
                '
   Screen.MousePointer = 11
   RNC$ = RECFILT$("COTELEF1", 3, MKI$(CICOTE%))
   URN& = ULTREG&("COTELEF1")
   '
   For U& = 1 To Len(RNC$) Step 4
      RENOTA& = CVS(Mid$(RNC$, U&, 4))
      If RENOTA& > 0 Then
         If RENOTA& <= URN& Then
            TTXX$ = REGLEIDO$("COTELEF1", RENOTA&)
            If Left$(TTXX$, 12) = NUCOTE$ Then
               GoTo 29
            End If
         End If
      End If
   Next U&
   TTXX$ = REGBLAN$("COTELEF1")
   '
29 AAA$ = Mid$(TTXX$, 179, 72)
   Screen.MousePointer = 1
   '
End Sub

Private Sub Text2_GotFocus()
   BOTREC.SetFocus
End Sub

Private Sub Text22_GotFocus()
   BOTREC.SetFocus
End Sub

Private Sub Text27_DblClick()
   Call Command15_Click
End Sub

Private Sub Text23_GotFocus()
   BOTREC.SetFocus
End Sub

Private Sub Text3_GotFocus()
   BOTREC.SetFocus
End Sub
