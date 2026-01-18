VERSION 5.00
Begin VB.Form EDIFORMU 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Documentos Electronicos"
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List1 
      Height          =   7665
      Left            =   2160
      TabIndex        =   26
      Top             =   120
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Command9"
      Height          =   615
      Left            =   1080
      TabIndex        =   25
      Top             =   240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   540
      Index           =   5
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   795
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   540
      Index           =   4
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   795
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   540
      Index           =   3
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   795
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   540
      Index           =   2
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   795
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   540
      Index           =   1
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   795
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   0
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   795
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Pantalla"
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
      Left            =   5670
      Picture         =   "EDIFORMU.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Pantalla Completa"
      Top             =   105
      Width           =   750
   End
   Begin VB.CommandButton Command7 
      Height          =   750
      Left            =   9555
      Picture         =   "EDIFORMU.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Pagina Siguiente"
      Top             =   105
      Width           =   450
   End
   Begin VB.CommandButton Command1 
      Height          =   750
      Left            =   9030
      Picture         =   "EDIFORMU.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Pagina Anterior"
      Top             =   105
      Width           =   450
   End
   Begin VB.CommandButton Command2 
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
      Height          =   750
      Left            =   10920
      Picture         =   "EDIFORMU.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Termina - Cierra la Aplicación"
      Top             =   105
      Width           =   750
   End
   Begin VB.CommandButton Command6 
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
      Height          =   750
      Left            =   7350
      Picture         =   "EDIFORMU.frx":0A68
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Ayuda Flexoft en Línea"
      Top             =   105
      Width           =   750
   End
   Begin VB.CommandButton Command3 
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   10080
      Picture         =   "EDIFORMU.frx":0D72
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Nuevo Documento"
      Top             =   105
      Width           =   765
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   8190
      Picture         =   "EDIFORMU.frx":107C
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Imprime Documento Seleccionado"
      Top             =   105
      Width           =   765
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   6510
      Picture         =   "EDIFORMU.frx":16E6
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Guarda Documento Grafico"
      Top             =   105
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   225
      LargeChange     =   100
      Left            =   0
      TabIndex        =   3
      Top             =   8050
      Width           =   11700
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   7070
      LargeChange     =   100
      Left            =   11700
      TabIndex        =   2
      Top             =   970
      Width           =   225
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   7050
      Left            =   0
      TabIndex        =   0
      Top             =   970
      Width           =   11670
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   7070
         Index           =   5
         Left            =   0
         ScaleHeight     =   7065
         ScaleWidth      =   11775
         TabIndex        =   13
         Top             =   0
         Visible         =   0   'False
         Width           =   11775
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   7070
         Index           =   4
         Left            =   420
         ScaleHeight     =   7065
         ScaleWidth      =   11700
         TabIndex        =   12
         Top             =   420
         Visible         =   0   'False
         Width           =   11700
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   7070
         Index           =   3
         Left            =   420
         ScaleHeight     =   7065
         ScaleWidth      =   11700
         TabIndex        =   11
         Top             =   420
         Visible         =   0   'False
         Width           =   11700
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   7070
         Index           =   2
         Left            =   315
         ScaleHeight     =   7065
         ScaleWidth      =   11700
         TabIndex        =   10
         Top             =   735
         Visible         =   0   'False
         Width           =   11700
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   7070
         Index           =   1
         Left            =   420
         ScaleHeight     =   7065
         ScaleWidth      =   11700
         TabIndex        =   9
         Top             =   420
         Visible         =   0   'False
         Width           =   11700
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   7070
         Index           =   0
         Left            =   420
         ScaleHeight     =   7065
         ScaleWidth      =   11700
         TabIndex        =   1
         Top             =   525
         Visible         =   0   'False
         Width           =   11700
      End
   End
   Begin VB.Label TEPRURED 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   195
      Left            =   2310
      TabIndex        =   24
      Top             =   420
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label TEPRUORI 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   105
      Left            =   2310
      TabIndex        =   23
      Top             =   105
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   11655
      Y1              =   8030
      Y2              =   8030
   End
   Begin VB.Line Line2 
      X1              =   11690
      X2              =   11690
      Y1              =   970
      Y2              =   8700
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   12000
      Y1              =   945
      Y2              =   945
   End
End
Attribute VB_Name = "EDIFORMU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   If PAGINACTIVA% > 0 Then
     Picture1(PAGINACTIVA%).Visible = False
     PAGINACTIVA% = PAGINACTIVA% - 1
     Picture1(PAGINACTIVA%).Visible = True
     Picture1(PAGINACTIVA%).Top = (-100) * VScroll1.Value
     Picture1(PAGINACTIVA%).Refresh
   End If
End Sub

Private Sub Command2_Click()
   Hide
End Sub

Private Sub Command3_Click()
   '
5  Call SELECHO("PDOCLST")
   DOCUNU& = Val(VALACT1$("PDOCLST"))
   If DOCUNU& < 1 Then
      Exit Sub
   End If
   REDOCU$ = REGACT$("PDOCLST")
   ORIPRINTER$ = TRIM$(Mid$(REDOCU$, 129))
   '
AORI$ = "P": If EXISTE%(LUDAT$ + "RDOCSPL.DT0") > 0 Then AORI$ = "R"
   Dim TD0 As String * 512
   NX0% = FILEOPEN%(LUDAT$ + AORI$ + "DOCSPL.DAT", 0, 512)
   Get #NX0%, 1, TD0$
   If CVS(Left$(TD0$, 4)) > DOCUNU& Or LOF(NX0%) < 1 Then
     Close #NX0%
     NX0% = FILEOPEN%(LUDAT$ + AORI$ + "DOCSPL.DT0", 0, 512)
   End If
   '
   URE& = Abs(LOF(NX0%) / 512)   ' PARA PREVENIR DESBORDAMIENTO DOSIVAC
   LI& = 0: LS& = URE&
6  E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
   L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
   '
   Get #NX0%, L&, TD0$
   DCC& = CVS(Left$(TD0$, 4))
   If DCC& < DOCUNU& Then LI& = L&: GoTo 6
   If DCC& > DOCUNU& Then LS& = L&: GoTo 6
   GoTo 8
   '
7  Call MENSERR(24, "Documento no Registrado")
   GoTo 5
   '
8  TPSP$ = ""
   Screen.MousePointer = 11
   While L& > 1
     Get #NX0%, L&, TD0$
     DCC& = CVS(Left$(TD0$, 4))
     If DCC& < DOCUNU& Then GoTo 9
     L& = L& - 1
   Wend
   '
9  CAPAGINAS% = 0
   For KKL& = L& To URE&
     Get #NX0%, KKL&, TD0$
     TBUF$ = TD0$
     If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
     If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
       TPSP$ = TPSP$ + Mid$(TBUF$, 5)
     End If
   Next KKL&
   '
95 SPOOLSTRING = TPSP$
   Close #NX0%
   For KKI% = 0 To 5
     Picture1(KKI%).Cls
     Picture1(KKI%).Top = 0
     Picture1(KKI%).Refresh
     Picture1(KKI%).Height = 7070
     Picture1(KKI%).Width = 11700
     Picture1(KKI%).DrawWidth = 3
     Picture1(KKI%).DrawStyle = 0
     Picture1(KKI%).FillStyle = 1
   Next KKI%
   ALTUPAGINA = 7070
   TOPEPAGINA = 0
   PAGINACTIVA% = 0
   CAPAGINAS% = 0
   '
   Screen.MousePointer = 11
10 PPXX% = InStr(TPSP$, Chr$(255))
   If PPXX% > 0 Then
      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
      TPSP$ = Mid$(TPSP$, PPXX% + 1)
      '
      If Len(COPRIAT$) > 0 Then
        COPRI% = Asc(Left$(COPRIAT$, 1))
        ATRIP$ = Mid$(COPRIAT$, 2)
        Call SHOWCOMMAND(COPRI%, ATRIP$)
      End If
      GoTo 10
      '
   End If
   '
   Screen.MousePointer = 1
   PAGINACTIVA% = 0
   ANCHPAGINA = EDIFORMU.Picture1(PAGINACTIVA%).Width
   EDIFORMU.Picture1(PAGINACTIVA%).Refresh
   EDIFORMU.VScroll1.Min = 0
   EDIFORMU.VScroll1.Value = 0
   EDIFORMU.VScroll1.Max = 1
   If ALTUPAGINA > EDIFORMU.Frame1.Height Then
      EDIFORMU.VScroll1.Max = (ALTUPAGINA - EDIFORMU.Frame1.Height) / 100
   End If
   EDIFORMU.HScroll1.Min = 0
   EDIFORMU.HScroll1.Value = 0
   EDIFORMU.HScroll1.Max = 1
   If ANCHPAGINA > EDIFORMU.Frame1.Width Then
      EDIFORMU.HScroll1.Max = (ANCHPAGINA - EDIFORMU.Frame1.Width) / 100
   End If
   For KKI% = 0 To CAPAGINAS%
      Picture1(KKI%).Visible = False
   Next KKI%
   Picture1(PAGINACTIVA%).Visible = True
   Picture1(PAGINACTIVA%).Refresh
   '
End Sub

Private Sub Command4_Click()
   '
   If DERACCE%("REPIMPRE", "Re-Imprimir Documentos Electronicos") < 2 Then
      Exit Sub
   End If
   '
   'Printer.DrawWidth = 3
   'Printer.DrawStyle = 0
   'Printer.FillStyle = 1
   '
   If SPOOLSTRING$ <> "" Then
      TPSP$ = SPOOLSTRING$
      GoTo 97
   End If
   '
   DOCUNU& = Val(VALACT1$("PDOCLST"))
   If DOCUNU& < 1 Then
      Exit Sub
   End If
   REDOCU$ = REGACT$("PDOCLST")
   ORIPRINTER$ = TRIM$(Mid$(REDOCU$, 129))
   URE& = ULTREG&("PDOCSPL")
   LI& = 0: LS& = URE&
6  E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
   L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
   DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
   If DCC& < DOCUNU& Then LI& = L&: GoTo 6
   If DCC& > DOCUNU& Then LS& = L&: GoTo 6
   GoTo 8
   '
7  Screen.MousePointer = 1
   Call MENSERR(24, "Documento no Registrado")
   Exit Sub
   '
8  TPSP$ = ""
   Screen.MousePointer = 11
   While L& > 1
     DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
     If DCC& < DOCUNU& Then GoTo 9
     L& = L& - 1
   Wend
   '
9  CAPAGINAS% = 0
   For KKL& = L& To URE&
     TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
     If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
     If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
       TPSP$ = TPSP$ + Mid$(TBUF$, 5)
     End If
   Next KKL&
   '
95 SPOOLSTRING$ = TPSP$
   '
97 Screen.MousePointer = 11
   On Error Resume Next
   MATRIXPUN% = 0
   If InStr(Printer.DeviceName, "810") > 0 Then MATRIXPUN% = 1
   If InStr(Printer.DeviceName, "286") > 0 Then MATRIXPUN% = 1
   If InStr(UCase$(Printer.DeviceName), "APEX") > 0 Then MATRIXPUN% = 1
   If InStr(UCase$(Printer.DeviceName), "FACTURA") > 0 Then MATRIXPUN% = 1
   If InStr(UCase$(Printer.DeviceName), "REMITO") > 0 Then MATRIXPUN% = 1
   'If InStr(Printer.DriverName, "EPSON9") > 0 Then MATRIXPUN% = 1
   '
   'Printer.EndDoc          ' INICIALIZA LA IMPRESORA
   '
   'If MATRIXPUN% = 1 Then
   '   Printer.PrintQuality = (-1)
   '   Printer.KillDoc
   'End If
   On Error GoTo 0
   '
   If Printer.DeviceName <> ORIPRINTER$ Then
      ORIPRINTER$ = REPLACAR$(ORIPRINTER$, "\", "/")
      MERRO$ = "ATENCION: Está Re-Imprimiendo en una Impresora Distinta"
      MERRO$ = MERRO$ + "\de la Originalmente Utilizada para este Documento."
      MERRO$ = MERRO$ + "\  \'" + ORIPRINTER$ + "'\  "
      MERRO$ = MERRO$ + "\  \Los Resultados de Impresión Pueden ser"
      MERRO$ = MERRO$ + "\Diferentes de los Originalmente Obtenidos."
      If COMALTER%(MERRO$ + "\\Continuar\Cancelar Reimpresión") <> 1 Then GoTo 99
   End If
   '
10 PPXX% = InStr(TPSP$, Chr$(255))
   If PPXX% > 0 Then
      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
      TPSP$ = Mid$(TPSP$, PPXX% + 1)
      '
      If Len(COPRIAT$) > 0 Then
        COPRI% = Asc(Left$(COPRIAT$, 1))
        ATRIP$ = Mid$(COPRIAT$, 2)
        Call PRINTCOMMAND(COPRI%, ATRIP$)
      End If
      GoTo 10
   End If
   '
99 'Printer.EndDoc
   Screen.MousePointer = 1
   '
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   Call HELPONLINE("EDIFORMU")
   Command6.Enabled = True
End Sub

Private Sub Command7_Click()
   If PAGINACTIVA% < CAPAGINAS% Then
     Picture1(PAGINACTIVA%).Visible = False
     PAGINACTIVA% = PAGINACTIVA% + 1
     Picture1(PAGINACTIVA%).Visible = True
     Picture1(PAGINACTIVA%).Top = (-100) * VScroll1.Value
     Picture1(PAGINACTIVA%).Refresh
   End If
End Sub

Private Sub Command8_Click()
'Clipboard.SetData Picture1(0).Image
   VScroll1.Visible = False
   HScroll1.Visible = False
   Command1.Enabled = False
   Command2.Enabled = False
   Command3.Enabled = False
   Command4.Enabled = False
   Command5.Enabled = False
   Command6.Enabled = False
   Command7.Enabled = False
   Command8.Enabled = False
   Line2.Visible = False
   '
   ALPAMAX = 0: ANPAMAX = 0
   For KKI% = 0 To 5
     If EDIFORMU.Picture1(KKI%).Width > ANPAMAX Then
       ANPAMAX = EDIFORMU.Picture1(KKI%).Width + 200
     End If
     If EDIFORMU.Picture1(KKI%).Height > ALPAMAX Then
       ALPAMAX = EDIFORMU.Picture1(KKI%).Height + 200
     End If
   Next KKI%
   REDUV = EDIFORMU.ScaleHeight / ALPAMAX
   REDUH = EDIFORMU.ScaleWidth / ANPAMAX
   REDUC = REDUV: If REDUH < REDUC Then REDUC = REDUH
   REDUFORMU = REDUC
   '
95 For KKI% = 0 To 5
     Picture2(KKI%).Cls
     Picture2(KKI%).Top = (EDIFORMU.ScaleHeight - ALPAMAX * REDUC) / 2
     Picture2(KKI%).Left = (EDIFORMU.ScaleWidth - ANPAMAX * REDUC) / 2
     Picture2(KKI%).Height = ALPAMAX * REDUC
     Picture2(KKI%).Width = ANPAMAX * REDUC
     Picture2(KKI%).Refresh
     Picture2(KKI%).DrawWidth = 2
     Picture2(KKI%).DrawStyle = 0
     Picture2(KKI%).FillStyle = 1
   Next KKI%
   '
   'ALTUPAGINA = 7070
   TOPEPAGINA = 0
   PAGINACTIVA% = 0
   CAPAGINAS% = 0
   '
   Screen.MousePointer = 11
   TPSP$ = SPOOLSTRING$
10 PPXX% = InStr(TPSP$, Chr$(255))
   If PPXX% > 0 Then
      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
      TPSP$ = Mid$(TPSP$, PPXX% + 1)
      '
      If Len(COPRIAT$) > 0 Then
        COPRI% = Asc(Left$(COPRIAT$, 1))
        ATRIP$ = Mid$(COPRIAT$, 2)
        Call SHOWCOMMZOOM(COPRI%, ATRIP$)
      End If
      GoTo 10
      '
   End If
   '
   Frame1.Visible = False
   Screen.MousePointer = 1
   PAGINACTIVA% = 0
   EDIFORMU.Picture2(PAGINACTIVA%).Refresh
   For KKI% = 0 To 5
      Picture1(KKI%).Visible = False
      Picture2(KKI%).Visible = False
   Next KKI%
   Picture2(PAGINACTIVA%).Visible = True
   Picture2(PAGINACTIVA%).Refresh
   '
End Sub



Private Sub Command9_Click()
5   ABUS$ = "Etiqueta de Cierre de O.F "   ' OF-008796-1"
   For U& = 1 To ULTREG&("PDOCLST")
      XX$ = REGLEIDO$("PDOCLST", U&)
      PPXX% = InStr(XX$, ABUS$)
      If PPXX% > 0 Then
        NOFA$ = TRIM$(Mid$(XX$, PPXX% + 26, 12))
        DOCUNU& = CVS(Left$(XX$, 4))
        GoSub 4
      End If
   Next U&
Exit Sub
Stop '
   '
4  AORI$ = "P": If EXISTE%(LUDAT$ + "RDOCSPL.DT0") > 0 Then AORI$ = "R"
   Dim TD0 As String * 512
   NX0% = FILEOPEN%(LUDAT$ + AORI$ + "DOCSPL.DAT", 0, 512)
   Get #NX0%, 1, TD0$
   If CVS(Left$(TD0$, 4)) > DOCUNU& Or LOF(NX0%) < 1 Then
     Close #NX0%
     NX0% = FILEOPEN%(LUDAT$ + AORI$ + "DOCSPL.DT0", 0, 512)
   End If
   '
   URE& = Abs(LOF(NX0%) / 512)   ' PARA PREVENIR DESBORDAMIENTO DOSIVAC
   LI& = 0: LS& = URE&
6  E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
   L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
   '
   Get #NX0%, L&, TD0$
   DCC& = CVS(Left$(TD0$, 4))
   If DCC& < DOCUNU& Then LI& = L&: GoTo 6
   If DCC& > DOCUNU& Then LS& = L&: GoTo 6
   GoTo 8
   '
7  ' Call MENSERR(24, "Documento no Registrado")
   GoTo 19
   '
8  TPSP$ = ""
   Screen.MousePointer = 11
   While L& > 1
     Get #NX0%, L&, TD0$
     DCC& = CVS(Left$(TD0$, 4))
     If DCC& < DOCUNU& Then GoTo 9
     L& = L& - 1
   Wend
   '
9  CAPAGINAS% = 0
   For KKL& = L& To URE&
     Get #NX0%, KKL&, TD0$
     TBUF$ = TD0$
     If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
     If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
       TPSP$ = TPSP$ + Mid$(TBUF$, 5)
     End If
   Next KKL&
   '
95 SPOOLSTRING = TPSP$
   Close #NX0%
   PPXX1% = InStr(TPSP$, NOFA$)
   If PPXX1% > 0 Then
      PPXX2% = InStr(1 + PPXX1%, TPSP$, NOFA$)
      If PPXX2% > PPXX1% Then
        PPXX3% = InStr(1 + PPXX2%, TPSP$, NOFA$)
        If PPXX3% > PPXX2% Then
          TTT$ = Mid$(TPSP$, PPXX3% - 16, 34)
List1.AddItem TTT$
'          MsgBox TTT$
        End If
      End If
   End If
   '
19 Return

End Sub

Private Sub Form_Load()
  EPAC$ = TRIM$(EMPREAC$)
  If EPAC$ <> "" Then
    Caption = Caption + "  -  " + EPAC$
  End If
  '
  Call CENTRAFORM(Me)
  For KKI% = 0 To 5
     Picture1(KKI%).Top = 0
     Picture1(KKI%).Left = 0
   Next KKI%
   Picture1(0).Visible = True
End Sub

Private Sub HScroll1_Change()
   Picture1(PAGINACTIVA%).Left = (-100) * HScroll1.Value
End Sub

Private Sub Picture2_DblClick(Index As Integer)
   For KKI% = 0 To 5
      Picture2(KKI%).Visible = False
   Next KKI%
   Frame1.Visible = True
   VScroll1.Visible = True
   HScroll1.Visible = True
   Command1.Enabled = True
   Command2.Enabled = True
   Command3.Enabled = True
   Command4.Enabled = True
   Command5.Enabled = True
   Command6.Enabled = True
   Command7.Enabled = True
   Command8.Enabled = True
   Line2.Visible = True
   '
   PAGINACTIVA% = 0
   EDIFORMU.Picture1(PAGINACTIVA%).Refresh
   Picture1(PAGINACTIVA%).Visible = True
   Picture1(PAGINACTIVA%).Refresh
   '
End Sub

Private Sub VScroll1_Change()
   Picture1(PAGINACTIVA%).Top = (-100) * VScroll1.Value
End Sub

Sub SHOWCOMMZOOM(COPRI%, ATRIP$)
   '
   If PAGINACTIVA% > CAPAGINAS% Then
      CAPAGINAS% = PAGINACTIVA%
   End If
10 On Error Resume Next
   Select Case COPRI%
   Case 1
     'Printer.Orientation = Val(TRIM$(ATRIP$))
   Case 2
     PAGINACTIVA% = PAGINACTIVA% + 1
   Case 3
     PAGINACTIVA% = PAGINACTIVA% + 1
   Case 4
     'ALTUPAGINA = Val(TRIM$(ATRIP$))
   Case 5
     'EDIFORMU.Picture1(PAGINACTIVA%).Width = Val(TRIM$(ATRIP$))
   Case 6
     EDIFORMU.Picture2(PAGINACTIVA%).ScaleMode = Val(TRIM$(ATRIP$))
   Case 7
     EDIFORMU.Picture2(PAGINACTIVA%).ForeColor = Val(TRIM$(ATRIP$))
   Case 8
     EDIFORMU.Picture2(PAGINACTIVA%).DrawMode = Val(TRIM$(ATRIP$))
   Case 9
     EDIFORMU.Picture2(PAGINACTIVA%).DrawWidth = 0.66 * Val(TRIM$(ATRIP$))
   Case 10
     EDIFORMU.Picture2(PAGINACTIVA%).DrawStyle = Val(TRIM$(ATRIP$))
   Case 11
     EDIFORMU.Picture2(PAGINACTIVA%).FillStyle = Val(TRIM$(ATRIP$))
   Case 12
     AA11 = Val(Mid$(ATRIP$, 1, 16)) * REDUFORMU
     BB11 = Val(Mid$(ATRIP$, 17, 16)) * REDUFORMU
     AA22 = Val(Mid$(ATRIP$, 33, 16)) * REDUFORMU
     BB22 = Val(Mid$(ATRIP$, 49, 16)) * REDUFORMU
     CCOLO = Val(Mid$(ATRIP$, 65, 16))
     'If BB11 + 1000 > ALTUPAGINA Then ALTUPAGINA = BB11 + 1000
     'If BB22 + 1000 > ALTUPAGINA Then ALTUPAGINA = BB22 + 1000
     'For KKI% = 0 To CAPAGINAS%
     '  If EDIFORMU.Picture1(KKI%).Height < ALTUPAGINA Then
     '    EDIFORMU.Picture1(KKI%).Height = ALTUPAGINA
     '  End If
     'Next KKI%
     'If AA11 > EDIFORMU.Picture1(PAGINACTIVA%).Width Then
     '   EDIFORMU.Picture1(PAGINACTIVA%).Width = AA11
     'End If
     'If AA22 > EDIFORMU.Picture1(PAGINACTIVA%).Width Then
     '   EDIFORMU.Picture1(PAGINACTIVA%).Width = AA22
     'End If
     BBLO$ = UCase$(TRIM$(Mid$(ATRIP$, 81, 16)))
     If BBLO$ = "B" Then
        EDIFORMU.Picture2(PAGINACTIVA%).Line (AA11, BB11 + TOPEPAGINA)-(AA22, BB22 + TOPEPAGINA), CCOLO, B
       Else
        EDIFORMU.Picture2(PAGINACTIVA%).Line (AA11, BB11 + TOPEPAGINA)-(AA22, BB22 + TOPEPAGINA), CCOLO
     End If
   Case 13
     EDIFORMU.TEPRUORI.FontName = TRIM$(ATRIP$)
   Case 14
     EDIFORMU.TEPRUORI.FontItalic = Val(TRIM$(ATRIP$))
   Case 15
     EDIFORMU.TEPRUORI.FontSize = Val(TRIM$(ATRIP$))
   Case 16
     EDIFORMU.TEPRUORI.FontBold = Val(TRIM$(ATRIP$))
   Case 17
     EDIFORMU.Picture2(PAGINACTIVA%).CurrentX = Val(TRIM$(ATRIP$)) * REDUFORMU
   Case 18
     EDIFORMU.Picture2(PAGINACTIVA%).CurrentY = (Val(TRIM$(ATRIP$)) + TOPEPAGINA) * REDUFORMU
   Case 19
     EDIFORMU.TEPRUORI.Caption = ATRIP$
     AMATEX = EDIFORMU.TEPRUORI.Width * REDUFORMU
     EDIFORMU.TEPRURED.Font = "Arial"
     If InStr(UCase$(EDIFORMU.TEPRUORI.FontName), "COURIER") > 0 Then
        EDIFORMU.TEPRURED.FontName = "Courier New"
     End If
     EDIFORMU.TEPRURED.FontBold = False
     EDIFORMU.TEPRURED.Caption = ATRIP$
     For KKKK = 24 To 1 Step -1
       EDIFORMU.TEPRURED.FontSize = KKKK / 2
       If EDIFORMU.TEPRURED.Width <= AMATEX Then GoTo 17
     Next KKKK
     '
17   EDIFORMU.Picture2(PAGINACTIVA%).FontName = EDIFORMU.TEPRURED.FontName '"Arial"
     EDIFORMU.Picture2(PAGINACTIVA%).FontSize = KKKK / 2
     EDIFORMU.Picture2(PAGINACTIVA%).FontBold = False
     ATRIP$ = RTrim$(ATRIP$)
     If Right$(ATRIP$, 1) = ";" Then
         ATRIP$ = Left$(ATRIP$, Len(ATRIP$) - 1)
         EDIFORMU.Picture2(PAGINACTIVA%).Print ATRIP$;
        Else
         EDIFORMU.Picture2(PAGINACTIVA%).Print ATRIP$
     End If
   Case 20
     EDIFORMU.Picture2(PAGINACTIVA%).FontUnderline = Val(TRIM$(ATRIP$))
   Case 21
     Call INCLUDEIMAGE(ATRIP$, "SHOW")
   End Select
   On Error GoTo 0
   '
End Sub
'

