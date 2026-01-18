VERSION 5.00
Begin VB.Form EDITFORM 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Documentos Electronicos"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command7 
      Height          =   750
      Left            =   9555
      Picture         =   "EDITFOR1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Pagina Siguiente"
      Top             =   105
      Width           =   450
   End
   Begin VB.CommandButton Command1 
      Height          =   750
      Left            =   9030
      Picture         =   "EDITFOR1.frx":030A
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
      Picture         =   "EDITFOR1.frx":0614
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
      Picture         =   "EDITFOR1.frx":075E
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
      Picture         =   "EDITFOR1.frx":0A68
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
      Picture         =   "EDITFOR1.frx":0D72
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
      Picture         =   "EDITFOR1.frx":13DC
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
         BackColor       =   &H80000009&
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
         BackColor       =   &H80000009&
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
         BackColor       =   &H80000009&
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
         BackColor       =   &H80000009&
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
         BackColor       =   &H80000009&
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
         BackColor       =   &H80000009&
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
      X2              =   11700
      Y1              =   945
      Y2              =   945
   End
End
Attribute VB_Name = "EDITFORM"
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
   '
   URE& = ULTREG&("PDOCSPL")
   LI& = 0: LS& = URE&
6  E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
   L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
   DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
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
   EDITFORM.Picture1(PAGINACTIVA%).Refresh
   EDITFORM.VScroll1.Min = 0
   EDITFORM.VScroll1.Value = 0
   EDITFORM.VScroll1.Max = 1
   If ALTUPAGINA > EDITFORM.Frame1.Height Then
      EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
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
   Printer.DrawWidth = 3
   Printer.DrawStyle = 0
   Printer.FillStyle = 1
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
   Printer.EndDoc          ' INICIALIZA LA IMPRESORA
   '
   If MATRIXPUN% = 1 Then
      Printer.PrintQuality = (-1)
      Printer.KillDoc
   End If
   On Error GoTo 0
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
99 Printer.EndDoc
   Screen.MousePointer = 1
   '
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   Call HELPONLINE("EDITFORM")
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
End Sub

Private Sub Form_Load()
  EPAC$ = TRIM$(EMPREAC$)
  If EPAC$ <> "" Then
      Caption = Caption + "  -  " + EPAC$
  End If
  '
  For KKI% = 0 To 5
     Picture1(KKI%).Top = 0
     Picture1(KKI%).Left = 0
   Next KKI%
   Picture1(0).Visible = True
End Sub

Private Sub HScroll1_Change()
   Picture1(PAGINACTIVA%).Left = (-100) * HScroll1.Value
End Sub

Private Sub VScroll1_Change()
   Picture1(PAGINACTIVA%).Top = (-100) * VScroll1.Value
End Sub
