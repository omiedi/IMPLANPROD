VERSION 5.00
Begin VB.Form AUTORISIM 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aprobación de Solicitudes Interna de Materiales "
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9510
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox LICOPEN 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   0
      ToolTipText     =   "Tilde las Solicitudes a Aprobar"
      Top             =   900
      Width           =   8325
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00E0E0E0&
      Height          =   435
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   8265
      TabIndex        =   8
      Top             =   495
      Width           =   8325
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "  Nro. Fecha    Sector       Observaciones    "
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   225
         Left            =   220
         TabIndex        =   9
         Top             =   105
         Width           =   10515
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00004000&
      Height          =   5775
      Left            =   8610
      ScaleHeight     =   5715
      ScaleWidth      =   1455
      TabIndex        =   3
      Top             =   0
      Width           =   1515
      Begin VB.CommandButton Command6 
         Caption         =   "Edit"
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
         Picture         =   "AUTORISIM.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Modificacion de la SIM"
         Top             =   2110
         Width           =   650
      End
      Begin VB.CommandButton Command4 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   120
         Picture         =   "AUTORISIM.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Ver por Pantalla Solicitud  Activa"
         Top             =   840
         Width           =   650
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   120
         Picture         =   "AUTORISIM.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Refresca - Invierte Selección"
         Top             =   1470
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   5
         Top             =   3720
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   6
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
         Height          =   650
         Left            =   120
         Picture         =   "AUTORISIM.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Aprueba la Selección"
         Top             =   3000
         Width           =   650
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
         Height          =   650
         Left            =   120
         Picture         =   "AUTORISIM.frx":0A68
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Cancela la Pre-Selección"
         Top             =   120
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "AUTORISIM.frx":0D72
         Top             =   3960
         Width           =   1515
      End
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   10200
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   5400
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label Label5 
      Height          =   435
      Left            =   7350
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label2 
      Height          =   435
      Left            =   7980
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Marque (tilde) las Solicitudes de Materiales que Desea Autorizar "
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   105
      TabIndex        =   11
      Top             =   105
      Width           =   6945
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   8450
      Y1              =   4270
      Y2              =   4270
   End
   Begin VB.Line Line2 
      X1              =   8445
      X2              =   8445
      Y1              =   525
      Y2              =   4320
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   330
      Left            =   7440
      TabIndex        =   10
      Top             =   1140
      Width           =   1695
   End
End
Attribute VB_Name = "AUTORISIM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BOTEXIT_Click()
   Label5 = "0"
   LICOPEN.Clear
   Unload Me
End Sub
'

Private Sub BOTREC_Click()
   '
   Label5 = "0"
   For I = 1 To LICOPEN.ListCount
     If LICOPEN.Selected(I - 1) = True Then
       CANSELE% = CANSELE% + 1
     End If
   Next I
   Label5 = TRIM$(Str$(CANSELE%))
   '
99 Label2 = "OK"
   '
   FAUTO = CVDAT(HOY(HOS$))
   For I = 1 To LICOPEN.ListCount
     If LICOPEN.Selected(I - 1) = True Then
       NUSOLI = XVALO(Left$(LICOPEN.List(I - 1), 6))
       USNU% = USNBR% Mod 100
       Call ACTUREGISTRO("ENCASOLINT", "AUTORIZA_1", "(AUTO_1 = " & USNU% & ") AND (AUTORIZA_1 IS NULL) AND NUME_SOLINT = " & NUSOLI, FAUTO, REGIS&)
       Call ACTUREGISTRO("ENCASOLINT", "AUTORIZA_2", "(AUTO_2 = " & USNU% & ") AND (AUTORIZA_2 IS NULL) AND NUME_SOLINT = " & NUSOLI, FAUTO, REGIS&)
       Call ACTUREGISTRO("ENCASOLINT", "AUTORIZA_3", "(AUTO_3 = " & USNU% & ") AND (AUTORIZA_3 IS NULL) AND NUME_SOLINT = " & NUSOLI, FAUTO, REGIS&)
       Call ACTUREGISTRO("ENCASOLINT", "AUTORIZA_4", "(AUTO_4 = " & USNU% & ") AND (AUTORIZA_4 IS NULL) AND NUME_SOLINT = " & NUSOLI, FAUTO, REGIS&)
     End If
   Next I
   '
   For I = 1 To LICOPEN.ListCount
     If LICOPEN.Selected(I - 1) = True Then
       TTXX$ = "Autoriza: " & USERNAME$ & " - " & TIMECENT$()
       Call ANOTASIM07.ANOTASIM(XVALO(Left$(LICOPEN.List(I - 1), 6)), TTXX$)
     End If
   Next I
   '
   Unload Me
   '
End Sub

Private Sub Command1_Click()
   command1.Enabled = False
   TPI = LICOPEN.TopIndex
   For U& = 1 To LICOPEN.ListCount
     If LICOPEN.Selected(U& - 1) = True Then
         LICOPEN.Selected(U& - 1) = False
       Else
         LICOPEN.Selected(U& - 1) = True
     End If
   Next U&
   On Error Resume Next
   LICOPEN.TopIndex = TPI
   On Error GoTo 0
   DoEvents
   command1.Enabled = True
End Sub

Private Sub Command4_Click()
    '
    Command4.Enabled = False
    ILISTA& = LICOPEN.ListIndex
    If ILISTA& >= O Then
      LISTACT$ = LICOPEN.TEXT
      NPX& = XVALO(Left$(LISTACT$, 6))
      If NPX& > 0 Then
        TIDOCUM$ = "Solicitud de Materiales"
        NUDOCU$ = TRIM$(Str$(NPX&))
        Screen.MousePointer = 11
        '
        For UI& = ULTREG&("PDOCLST") To 1 Step -1
             XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
             If InStr(XX$, TIDOCUM$) > 4 Then
               PPXX% = InStr(XX$, "Nro.")
               If PPXX% > 0 Then
                 XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
               End If
               If InStr(XX$, NUDOCU$) > 4 Then
                 DOCUNU& = CVS(Left$(XX$, 4))
                 GoTo 5
               End If
             End If
        Next UI&
           '
        Screen.MousePointer = 1
        Call COMUNI("No se Encontró el Documento Impreso\Correspondiente a Esta Solicitud de Material.\ ")
           '
5       Call MUESTRADOC(DOCUNU&) 'PERMITE VER LOS DOCUMENTOS ANTIGUOS POR PANTALLA
      End If
    End If
99  Command4.Enabled = True

         '
'         For UI& = ULTREG&("PDOCLST") To 1 Step -1
'           XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
'           If InStr(XX$, TIDOCUM$) > 4 Then
'             PPXX% = InStr(XX$, "Nro.")
'             If PPXX% > 0 Then
'               XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
'             End If
'             If InStr(XX$, NUDOCU1$) > 4 Or InStr(XX$, NUDOCU2$) > 4 Then
'               DOCUNU& = CVS(Left$(XX$, 4))
'               GoTo 5
'             End If
'           End If
'         Next UI&
'         Screen.MousePointer = 1
'         Call COMUNI("Documento no Encontrado")
'         GoTo 99
         '
'5      Call MUESTRADOC(DOCUNU&) 'PERMITE VER LOS DOCUMENTOS ANTIGUOS POR PANTALLA
'5         For KKI% = 0 To 5
'           EDITFORM.Picture1(KKI%).Cls
'           EDITFORM.Picture1(KKI%).DrawWidth = 3
'           EDITFORM.Picture1(KKI%).DrawStyle = 0
'           EDITFORM.Picture1(KKI%).FillStyle = 1
'         Next KKI%
'         '
'         REDOCU$ = REGACT$("PDOCLST")
'         URE& = ULTREG&("PDOCSPL")
'         LI& = 0: LS& = URE&
'6         E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
'         L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
'         DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'         If DCC& < DOCUNU& Then LI& = L&: GoTo 6
'         If DCC& > DOCUNU& Then LS& = L&: GoTo 6
'         GoTo 8
'         '
'7         Screen.MousePointer = 1
'         Call MENSERR(24, "Documento no Registrado")
'         GoTo 99
'         '
'8         TPSP$ = ""
'         Screen.MousePointer = 11
'         While L& > 1
'           DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'           If DCC& < DOCUNU& Then GoTo 9
'           L& = L& - 1
'         Wend
'         '
'9         CAPAGINAS% = 0
'         For KKL& = L& To URE&
'           TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
'           If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
'           If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
'             TPSP$ = TPSP$ + Mid$(TBUF$, 5)
'           End If
'         Next KKL&
'         '
'95        SPOOLSTRING$ = TPSP$
'         FIN& = Len(TPSP$)
'         For KKI% = 0 To 5
'           EDITFORM.Picture1(KKI%).Cls
'           EDITFORM.Picture1(KKI%).Top = 0
'           EDITFORM.Picture1(KKI%).Refresh
'           EDITFORM.Picture1(KKI%).Height = 7070
'           EDITFORM.Picture1(KKI%).Width = 11700
'           EDITFORM.Picture1(KKI%).DrawWidth = 3
'           EDITFORM.Picture1(KKI%).DrawStyle = 0
'           EDITFORM.Picture1(KKI%).FillStyle = 1
'         Next KKI%
'         ALTUPAGINA = 7070
'         TOPEPAGINA = 0
'         PAGINACTIVA% = 0
'         CAPAGINAS% = 0
'         '
'10        PPXX% = InStr(TPSP$, Chr$(255))
'         If PPXX% > 0 Then
'           COPRIAT$ = Left$(TPSP$, PPXX% - 1)
'           TPSP$ = Mid$(TPSP$, PPXX% + 1)
'           Call TRACE(20, FIN& - Len(TPSP$), FIN&)
'           '
'           If Len(COPRIAT$) > 0 Then
'             COPRI% = Asc(Left$(COPRIAT$, 1))
'             ATRIP$ = Mid$(COPRIAT$, 2)
'             Call SHOWCOMMAND(COPRI%, ATRIP$)
'           End If
'           GoTo 10
'           '
'         End If
'         '
'         UI& = UI& + 1
'         XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
'         If InStr(XX$, TIDOCUM$) > 4 Then
'           If InStr(XX$, NUDOCU$) > 4 Then
'             DOCUNU& = CVS(Left$(XX$, 4))
'             GoTo 7
'           End If
'         End If
'         '
'         Screen.MousePointer = 1
'         PAGINACTIVA% = 0
'         For KKI% = 0 To 5
'            EDITFORM.Picture1(KKI%).Visible = False
'         Next KKI%
'         EDITFORM.Picture1(PAGINACTIVA%).Visible = True
'         EDITFORM.Picture1(PAGINACTIVA%).Refresh
'         EDITFORM.VScroll1.Min = 0
'         EDITFORM.VScroll1.Value = 0
'         EDITFORM.VScroll1.Max = 1
'         If ALTUPAGINA > EDITFORM.Frame1.Height Then
'            EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
'         End If
'         ANCHPAGINA = EDITFORM.Picture1(PAGINACTIVA%).Width
'         EDITFORM.HScroll1.Min = 0
'         EDITFORM.HScroll1.Value = 0
'         EDITFORM.HScroll1.Max = 1
'         If ANCHPAGINA > EDITFORM.Frame1.Width Then
'           EDITFORM.HScroll1.Max = (ANCHPAGINA - EDITFORM.Frame1.Width) / 100
'         End If
'         EDITFORM.Command3.Enabled = False
'         EDITFORM.Command5.Enabled = False
'         EDITFORM.Show 1
'         '
    '
End Sub

Private Sub Command6_Click()
    If DERACCE%("MODISIM", "Modificación de S.I.M.s") > 1 Then
      MODISIM07.Show 1
    End If
    '
End Sub

Private Sub Form_Load()
    '
    USNU% = USNBR% Mod 100
    '
    CONDI$ = "(Status <> 9 OR STATUS IS NULL) "                                      'SIM's NO ANULADAS
    CONDI$ = CONDI$ + " AND ((AUTO_1 = " & USNU% & " AND AUTORIZA_1 IS NULL)"
    CONDI$ = CONDI$ + " OR (AUTO_2 = " & USNU% & " And (AUTORIZA_2 Is Null OR AUTORIZA_2 < 1) AND AUTORIZA_1 IS NOT NULL AND AUTORIZA_1 > 0)"
    CONDI$ = CONDI$ + " OR (AUTO_3 = " & USNU% & " And (AUTORIZA_3 Is Null OR AUTORIZA_3 < 1) AND AUTORIZA_2 IS NOT NULL AND AUTORIZA_2 > 0)"
    CONDI$ = CONDI$ + " OR (AUTO_4 = " & USNU% & " And (AUTORIZA_4 Is Null OR AUTORIZA_4 < 1) AND AUTORIZA_3 IS NOT NULL AND AUTORIZA_3 > 0))"
    CONDI$ = CONDI$ + " ORDER BY NUME_SOLINT ASC"
    '
    Call CARGALISTA(LICOPEN, "ENCASOLINT", " NUME_SOLINT/F6; FEMI_SOLINT/D8; SECT_SOLINT/A12; OBSE_SOLINT/A24 ", CONDI$)

End Sub

