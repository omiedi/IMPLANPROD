VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form EDITFORM 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00DAE4EB&
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
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command19 
      Caption         =   "Genera Registro en Pdoclst"
      Height          =   255
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Genera Registro en Pdoclst"
      Height          =   255
      Left            =   0
      TabIndex        =   34
      Top             =   600
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   630
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   30
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000005&
      Height          =   540
      Index           =   5
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   795
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000005&
      Height          =   540
      Index           =   4
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   795
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000005&
      Height          =   540
      Index           =   3
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   795
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000005&
      Height          =   540
      Index           =   2
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   795
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000005&
      Height          =   540
      Index           =   1
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   795
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
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
      Left            =   -120
      ScaleHeight     =   480
      ScaleWidth      =   795
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DAE4EB&
      Caption         =   "IMPRESORA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3120
      TabIndex        =   25
      Top             =   0
      Width           =   4005
      Begin VB.CommandButton Command4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3440
         Picture         =   "EDITFORM.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Imprime Documento Seleccionado"
         Top             =   290
         Width           =   450
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   210
         TabIndex        =   26
         Text            =   "Combo1"
         Top             =   380
         Width           =   3045
      End
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   770
      Left            =   8630
      Picture         =   "EDITFORM.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Pagina Anterior"
      Top             =   120
      Width           =   450
   End
   Begin VB.CommandButton Command3 
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   770
      Left            =   9565
      Picture         =   "EDITFORM.frx":0974
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Nuevo Documento"
      Top             =   120
      Width           =   660
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   770
      Left            =   10195
      Picture         =   "EDITFORM.frx":0C7E
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Ayuda Flexoft en Línea"
      Top             =   120
      Width           =   645
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Scr"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   770
      Left            =   7980
      Picture         =   "EDITFORM.frx":0F88
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Pantalla Completa"
      Top             =   120
      Width           =   645
   End
   Begin VB.CommandButton Command7 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   770
      Left            =   9080
      Picture         =   "EDITFORM.frx":13CA
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Pagina Siguiente"
      Top             =   120
      Width           =   450
   End
   Begin VB.CommandButton Command2 
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
      Height          =   770
      Left            =   10825
      Picture         =   "EDITFORM.frx":16D4
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Termina - Cierra la Aplicación"
      Top             =   120
      Width           =   645
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
      Top             =   960
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
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   7070
         Index           =   5
         Left            =   0
         ScaleHeight     =   7065
         ScaleWidth      =   11775
         TabIndex        =   11
         Top             =   0
         Visible         =   0   'False
         Width           =   11775
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   7070
         Index           =   4
         Left            =   420
         ScaleHeight     =   7065
         ScaleWidth      =   11700
         TabIndex        =   10
         Top             =   420
         Visible         =   0   'False
         Width           =   11700
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   7070
         Index           =   3
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
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   7070
         Index           =   2
         Left            =   315
         ScaleHeight     =   7065
         ScaleWidth      =   11700
         TabIndex        =   8
         Top             =   735
         Visible         =   0   'False
         Width           =   11700
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   7070
         Index           =   1
         Left            =   420
         ScaleHeight     =   7065
         ScaleWidth      =   11700
         TabIndex        =   7
         Top             =   420
         Visible         =   0   'False
         Width           =   11700
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
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
   Begin VB.CommandButton Command12 
      Caption         =   "Mail"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   770
      Left            =   7320
      Picture         =   "EDITFORM.frx":181E
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Envia Documento por Mail"
      Top             =   120
      Width           =   645
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "EDITFORM.frx":1B28
      TabIndex        =   31
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "EDITFORM.frx":1B86
      Top             =   0
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   4200
      TabIndex        =   33
      Top             =   120
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label LblTipodeDoc 
      Height          =   255
      Left            =   0
      TabIndex        =   32
      Top             =   720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label98 
      Height          =   255
      Left            =   1320
      TabIndex        =   29
      Top             =   360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label2 
      Height          =   450
      Left            =   2520
      TabIndex        =   24
      Top             =   525
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label Label1 
      Height          =   330
      Left            =   2625
      TabIndex        =   23
      Top             =   -105
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label TEPRURED 
      AutoSize        =   -1  'True
      Caption         =   "teprured"
      Height          =   195
      Left            =   2310
      TabIndex        =   22
      Top             =   420
      Visible         =   0   'False
      Width           =   900
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
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   105
      Left            =   2310
      TabIndex        =   21
      Top             =   105
      Visible         =   0   'False
      Width           =   330
   End
End
Attribute VB_Name = "EDITFORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim REDUFORMU

Private Sub Command1_Click()
   If PAGINACTIVA% > 0 Then
     Picture1(PAGINACTIVA%).Visible = False
     PAGINACTIVA% = PAGINACTIVA% - 1
     Picture1(PAGINACTIVA%).Visible = True
     Picture1(PAGINACTIVA%).Top = (-100) * VScroll1.Value
     Picture1(PAGINACTIVA%).Refresh
   End If
End Sub

Private Sub Command10_Click()
   Call SELECHO("DEUDOR1")
   If VALACT1$("DEUDOR1") <> "" Then
     NC% = XVALO(VALACT1$("DEUDOR1"))
     Text3 = ""
     Text2 = rasocli$(NC%)
     Text12 = CONTACTCLI$(NC%)
     Text4 = EMAILCLI$(NC%)
     Text5 = "Estamos enviando adjunto "
   End If
End Sub

Private Sub Command12_Click()
   Command12.Enabled = False
   On Error Resume Next
   FMAIL09.Show 1
   If Err Then
      On Error GoTo 0
      Call COMUNI("Imposible Enviar por Mail\en la Presente Versión.\   \Solicite Actualización de Versión.")
   End If
   On Error GoTo 0
   Command12.Enabled = True
End Sub

Private Sub Command14_Click()
   If SPOOLSTRING$ <> "" Then
     TPSP$ = SPOOLSTRING$
     SENDBYMAIL$ = Text4
     '
     Label1 = "DOCADJUN"
   End If
End Sub

Private Sub Command19_Click()
'For i = 1 To 2240
'    Call Command5_Click
'  Next i
End Sub

Private Sub Command2_Click()
   Hide
End Sub

Private Sub Command22_Click()
   Call Command9_Click
End Sub

Private Sub Command3_Click()
   '
5  Call SELECHO("PDOCLST")
   DOCUNU& = XVALO(VALACT1$("PDOCLST"))
   If DOCUNU& < 1 Then
      Exit Sub
   End If
   REDOCU$ = REGACT$("PDOCLST")
   ORIPRINTER$ = TRIM$(Mid$(REDOCU$, 129))
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
   NUPIC% = EDITFORM.Picture1.Count
   For KKI% = 0 To NUPIC% - 1
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

Sub Command4_Click()
   '
   If DERACCE%("REPIMPRE", "Re-Imprimir Documentos Electronicos") < 2 Then
      Exit Sub
   End If
   '
   If SPOOLSTRING$ <> "" Then
      TPSP$ = SPOOLSTRING$
      GoTo 97
   End If
   '
   DOCUNU& = XVALO(VALACT1$("PDOCLST"))
   If DOCUNU& < 1 Then
      Exit Sub
   End If
   REDOCU$ = REGACT$("PDOCLST")
   ORIPRINTER$ = TRIM$(Mid$(REDOCU$, 129))
   Call MUESTRADOC(DOCUNU&, "PRINT")
   Exit Sub
   '
97 Screen.MousePointer = 11
   PRINTERPORT$ = Combo1.TEXT
   On Error Resume Next
   MATRIXPUN% = 0
   If InStr(Printer.DeviceName, "810") > 0 Then MATRIXPUN% = 1
   If InStr(Printer.DeviceName, "286") > 0 Then MATRIXPUN% = 1
   If InStr(UCase$(Printer.DeviceName), "APEX") > 0 Then MATRIXPUN% = 1
   If InStr(UCase$(Printer.DeviceName), "FACTURA") > 0 Then MATRIXPUN% = 1
   If InStr(UCase$(Printer.DeviceName), "REMITO") > 0 Then MATRIXPUN% = 1
   On Error GoTo 0
   '
   On Error Resume Next
   PDVN$ = Printer.DeviceName
   On Error GoTo 0
   If PDVN$ <> "" Then
     If PDVN$ <> ORIPRINTER$ Then
       If TRIM$(ORIPRINTER$) <> "" Then
         ORIPRIN1$ = REPLACAR$(ORIPRINTER$, "\", "/")
         MERRO$ = "ATENCION: Está Re-Imprimiendo en una Impresora Distinta"
         MERRO$ = MERRO$ + "\de la Originalmente Utilizada para este Documento."
         MERRO$ = MERRO$ + "\  \'" + ORIPRIN1$ + "'\  "
         MERRO$ = MERRO$ + "\  \Los Resultados de Impresión Pueden ser"
         MERRO$ = MERRO$ + "\Diferentes de los Originalmente Obtenidos."
         If COMALTER%(MERRO$ + "\\Continuar\Cancelar Reimpresión") <> 1 Then GoTo 99
       End If
     End If
   End If
   '
   If Combo1.TEXT <> PDVN$ Then
     Printer.TrackDefault = False
     Call SETPRINPORT(Combo1.TEXT)
     PRINTERPORT$ = Combo1.TEXT
   End If
   '
   FUERADEAREA% = 0
   FUENODISPO% = 0
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
      If InStr(UCase$(PRINTERPORT$), "PDF") > 0 Then
         If COPRI% = 3 Then    ' CORRESPONDE A ENDDOC
           GoTo 15
         End If
      End If
      GoTo 10
   End If
   '
15 If FUERADEAREA% > 0 Then
      Call MENSERR(24, "MENSAJE CRÍTICO: Se Han Encontrado Leyendas\Fuera de Area de Impresión Horizontal o Vertical.\   \FLEXOFT ha Suprimido estas Leyendas.\  \Se Recomienda Revisar y Corregir el Tamaño de Papel\Configurado en Esta Estación de Trabajo\para la Impresora " + REPLACAR$(Printer.DeviceName, "\", "/") + ".")
   End If
   FUERADEAREA% = 0
   '
   If FUENODISPO% > 0 Then
      Call MENSERR(24, "MENSAJE CRÍTICO: El Formulario Electrónico Actual\Utiliza Fuentes que no Están Disponibles\en la Impresora " + REPLACAR$(Printer.DeviceName, "\", "/") + ".\   \Se Recomienda Revisar el Formulario Electrónico\o Utilizar Otra Impresora de la Red.")
   End If
   FUENODISPO% = 0
   '
99 If Combo1.TEXT <> PDVN$ Then
     Call SETPRINPORT("RESTORE")
     PRINTERPORT$ = ""
     On Error Resume Next
     Combo1.TEXT = Printer.DeviceName
     On Error GoTo 0
   End If
   Screen.MousePointer = 1
   '
End Sub

 Sub Command5_Click()
'Este proceso se creo para generar nuevos registros en pdoclst, en Eym luego de un problema se recuperaron los datos
'pero en sql quedaron los datos correctos pero el pdoclst se recupero un archivo de un mes anterior
'por tal motivo y como toma el numero siguiente del pdoclst (ultreg) generaba el registro en sql duplicado por que el mismo ya existia
'ejemplo el pdoclst tenia 5000 registros y el pdoc_tbl de sql tenia 5100 por lo tanto empezo a duplicar desde el 5000 en mas.
'accion que se tomo generar tantos registros para empatar el sql
'modificar manualmente en la tabla sql los duplicados con el numero que le hubiera correspondido.

20 Call BLOQARCH("PDOCLST")
   If Err Then
21    On Error GoTo 0
      Call COMUNI("Error al Bloquear Tablas'PDOCLST'.\  \Espere un Instante y Pulse 'Aceptar'\para Reanudar Ejecución.")
      GoTo 20
   End If
   On Error GoTo 0
   '
   UDOCU& = 0
   UREGI& = ULTREG&("PDOCLST")
   If UREGI& > 0 Then
      UDOCU& = CVS(Left$(REGLEIDO$("PDOCLST", UREGI&), 4))
   End If
   PDLS$ = String$(256, 0)
   DOCUNU& = 1 + UDOCU&
   HOII% = HOY(HOS$)
   HORX% = Int(HORANUM(Time$))
   Call REPLA(PDLS$, MKS$(DOCUNU&), 1, 4)
   Call REPLA(PDLS$, "RECUPERO", 5, 60)
   Call REPLA(PDLS$, MKI$(HOII%), 65, 2)
   Call REPLA(PDLS$, MKI$(HORX%), 67, 2)
   Call REPLA(PDLS$, USERNAME$, 69, 30)
   Call REPLA(PDLS$, NOMTER$, 99, 30)
     On Error Resume Next
   Call REPLA(PDLS$, Printer.DeviceName, 129, 128)
     On Error GoTo 0
   Call REGAPP("PDOCLST", PDLS$)
   Call CIERRARCH("PDOCLST")
   Command5.Caption = UREGI&
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
     On Error Resume Next
     Picture1(PAGINACTIVA%).Visible = True
     If Err Then
       On Error GoTo 0
       Call MENSERR(24, "Imposible Mostrar Página " + TRIM$(Str$(PAGINACTIVA% + 1)))
       PAGINACTIVA% = PAGINACTIVA% - 1
       Exit Sub
     End If
     On Error GoTo 0
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
   Command6.Enabled = False
   Command7.Enabled = False
   Command8.Enabled = False
   Command12.Enabled = False
   Line2.Visible = False
   '
   ALPAMAX = 0: ANPAMAX = 0
   NUPIC% = EDITFORM.Picture1.Count
   For KKI% = 0 To NUPIC% - 1
     If EDITFORM.Picture1(KKI%).Width > ANPAMAX Then
       ANPAMAX = EDITFORM.Picture1(KKI%).Width + 200
     End If
     If EDITFORM.Picture1(KKI%).Height > ALPAMAX Then
       ALPAMAX = EDITFORM.Picture1(KKI%).Height + 200
     End If
   Next KKI%
   REDUV = EDITFORM.ScaleHeight / ALPAMAX
   REDUH = EDITFORM.ScaleWidth / ANPAMAX
   REDUC = REDUV: If REDUH < REDUC Then REDUC = REDUH
   REDUFORMU = REDUC
   '
95 NUPIC% = EDITFORM.Picture1.Count
   For KKI% = 0 To NUPIC% - 1
     Picture2(KKI%).Cls
     Picture2(KKI%).Top = (EDITFORM.ScaleHeight - ALPAMAX * REDUC) / 2
     Picture2(KKI%).Left = (EDITFORM.ScaleWidth - ANPAMAX * REDUC) / 2
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
   EDITFORM.Picture2(PAGINACTIVA%).Refresh
   NUPIC% = EDITFORM.Picture1.Count
   For KKI% = 0 To NUPIC% - 1
      Picture1(KKI%).Visible = False
      Picture2(KKI%).Visible = False
   Next KKI%
   Picture2(PAGINACTIVA%).Visible = True
   Picture2(PAGINACTIVA%).Refresh
   '
End Sub

Private Sub Command9_Click()
  Text1 = ""
  Text2 = ""
  Text3 = ""
  Text4 = ""
  Text5 = ""
  Text12 = ""
  Visible = False
End Sub

Private Sub Form_Load()
  '
   If Control("NEWASPEC", "CAMBIOAS") = "SI" Then
      Me.BackColor = &HFCD7B6
      Me.Frame1.BackColor = &HFCD7B6
      Me.Frame2.BackColor = &HFCD7B6
   End If

  Call TRANSLABELS(Name)
  Call CENTRAFORM(Me)
  '
  UTILIZA_SKIN% = 1
  Call APLICACIONSKINS(Me)

  EPAC$ = TRIM$(EMPREAC$)
  If EPAC$ <> "" Then
       EMPRE_SIN_USUARIO$ = REPLACAR$(EPAC$, USERNAME$, "")
       Caption = Caption + "  -  " + EMPRE_SIN_USUARIO$
  End If
    '
  NUPIC% = EDITFORM.Picture1.Count
  For KKI% = 0 To NUPIC% - 1
    Picture1(KKI%).Top = 0
    Picture1(KKI%).Left = 0
  Next KKI%
  Picture1(0).Visible = True
  If TRIM$(UCase$(App.EXEName)) <> "PRIFORMS" Then
    Command3.Enabled = False
  End If
  '
  Combo1.Clear
  Dim PRX As Printer
  For Each PRX In Printers
    Combo1.AddItem PRX.DeviceName
  Next
  On Error Resume Next
  Combo1.TEXT = Printer.DeviceName
  On Error GoTo 0
'  Command5.Enabled = True
  '
  
End Sub

Private Sub HScroll1_Change()
   On Error Resume Next
   Picture1(PAGINACTIVA%).Left = (-100) * HScroll1.Value
   On Error GoTo 0
End Sub

Private Sub HScroll1_Scroll()
   On Error Resume Next
   Picture1(PAGINACTIVA%).Left = (-100) * HScroll1.Value
   On Error GoTo 0
End Sub

Private Sub Picture2_DblClick(Index As Integer)
   NUPIC% = EDITFORM.Picture1.Count
   For KKI% = 0 To NUPIC% - 1
      Picture2(KKI%).Visible = False
   Next KKI%
   Frame1.Visible = True
   VScroll1.Visible = True
   HScroll1.Visible = True
   Command1.Enabled = True
   Command2.Enabled = True
   If TRIM$(UCase$(App.EXEName)) = "PRIFORMS" Then
     Command3.Enabled = True
   End If
   Command4.Enabled = True
   Command6.Enabled = True
   Command7.Enabled = True
   Command8.Enabled = True
   Command12.Enabled = True
   Line2.Visible = True
   '
   PAGINACTIVA% = 0
   EDITFORM.Picture1(PAGINACTIVA%).Refresh
   Picture1(PAGINACTIVA%).Visible = True
   Picture1(PAGINACTIVA%).Refresh
   '
End Sub

Private Sub Timer1_Timer()
   NUCETE$ = TRIM$(Label1)
   If NUCETE$ = "" Then
     Exit Sub
   End If
   '
   LETR$ = "N"
   On Error Resume Next
   AppActivate "Save PDF"
   If Err Then
     LETR$ = "O"
     On Error GoTo 0
     On Error Resume Next
     AppActivate "Guardar como"
     If Err Then
        On Error GoTo 0
        Exit Sub
     End If
   End If
   On Error GoTo 0
   '
   'Se modifica por que en la ventana de windows seven de guardar como la letra es 'O' en vez de N
   If LETR$ = "N" Then SendKeys "%N" + LUCOM$ + "" + NUCETE$ + ".PDF" + Chr$(13)

   If LETR$ = "O" Then SendKeys "%o" + LUCOM$ + "" + NUCETE$ + ".PDF" + Chr$(13)
   
'   LETR$ = "O"
'   On Error Resume Next
'   AppActivate "Guardar como"
'   If Err Then
'     LETR$ = "N"
'     On Error GoTo 0
'     On Error Resume Next
'     AppActivate "Save PDF"
'     If Err Then
'        On Error GoTo 0
'        Exit Sub
'     End If
'   End If
'   On Error GoTo 0
'   '
''   RUXA$ = CurDir
'   'Se modifica por que en la ventana de windows seven de guardar como la letra es 'O' en vez de N
'   If LETR$ = "N" Then SendKeys "%N" + LUCOM$ + "" + NUCETE$ + ".PDF" + Chr$(13)
'
'   If LETR$ = "O" Then SendKeys "%o" + LUCOM$ + "" + NUCETE$ + ".PDF" + Chr$(13) ' ANTES ERA O AHORA CAMBIO A N
   NUCETE$ = ""
   Timer1.Enabled = False
'   ChDir RUXA$
   Exit Sub
   '
End Sub
'
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
     'Printer.Orientation = XVALO(TRIM$(ATRIP$))
   Case 2
     PAGINACTIVA% = PAGINACTIVA% + 1
   Case 3
     PAGINACTIVA% = PAGINACTIVA% + 1
   Case 4
     'ALTUPAGINA = XVALO(TRIM$(ATRIP$))
   Case 5
     'EDITFORM.Picture1(PAGINACTIVA%).Width = XVALO(TRIM$(ATRIP$))
   Case 6
     EDITFORM.Picture2(PAGINACTIVA%).ScaleMode = XVALO(TRIM$(ATRIP$))
   Case 7
     EDITFORM.Picture2(PAGINACTIVA%).ForeColor = XVALO(TRIM$(ATRIP$))
   Case 8
     EDITFORM.Picture2(PAGINACTIVA%).DrawMode = XVALO(TRIM$(ATRIP$))
   Case 9
     EDITFORM.Picture2(PAGINACTIVA%).DrawWidth = 0.66 * XVALO(TRIM$(ATRIP$))
   Case 10
     EDITFORM.Picture2(PAGINACTIVA%).DrawStyle = XVALO(TRIM$(ATRIP$))
   Case 11
     EDITFORM.Picture2(PAGINACTIVA%).FillStyle = XVALO(TRIM$(ATRIP$))
   Case 12
     AA11 = XVALO(Mid$(ATRIP$, 1, 16)) * REDUFORMU
     BB11 = XVALO(Mid$(ATRIP$, 17, 16)) * REDUFORMU
     AA22 = XVALO(Mid$(ATRIP$, 33, 16)) * REDUFORMU
     BB22 = XVALO(Mid$(ATRIP$, 49, 16)) * REDUFORMU
     CCOLO = XVALO(Mid$(ATRIP$, 65, 16))
     'If BB11 + 1000 > ALTUPAGINA Then ALTUPAGINA = BB11 + 1000
     'If BB22 + 1000 > ALTUPAGINA Then ALTUPAGINA = BB22 + 1000
     'For KKI% = 0 To CAPAGINAS%
     '  If EDITFORM.Picture1(KKI%).Height < ALTUPAGINA Then
     '    EDITFORM.Picture1(KKI%).Height = ALTUPAGINA
     '  End If
     'Next KKI%
     'If AA11 > EDITFORM.Picture1(PAGINACTIVA%).Width Then
     '   EDITFORM.Picture1(PAGINACTIVA%).Width = AA11
     'End If
     'If AA22 > EDITFORM.Picture1(PAGINACTIVA%).Width Then
     '   EDITFORM.Picture1(PAGINACTIVA%).Width = AA22
     'End If
     BBLO$ = UCase$(TRIM$(Mid$(ATRIP$, 81, 16)))
     If BBLO$ = "B" Then
        EDITFORM.Picture2(PAGINACTIVA%).Line (AA11, BB11 + TOPEPAGINA)-(AA22, BB22 + TOPEPAGINA), CCOLO, B
       Else
        EDITFORM.Picture2(PAGINACTIVA%).Line (AA11, BB11 + TOPEPAGINA)-(AA22, BB22 + TOPEPAGINA), CCOLO
     End If
   Case 13
     EDITFORM.TEPRUORI.FontName = TRIM$(ATRIP$)
   Case 14
     EDITFORM.TEPRUORI.FontItalic = XVALO(TRIM$(ATRIP$))
   Case 15
     EDITFORM.TEPRUORI.FontSize = XVALO(TRIM$(ATRIP$))
   Case 16
     EDITFORM.TEPRUORI.FontBold = XVALO(TRIM$(ATRIP$))
   Case 17
     EDITFORM.Picture2(PAGINACTIVA%).CurrentX = XVALO(TRIM$(ATRIP$)) * REDUFORMU
   Case 18
     EDITFORM.Picture2(PAGINACTIVA%).CurrentY = (XVALO(TRIM$(ATRIP$)) + TOPEPAGINA) * REDUFORMU
   Case 19
     EDITFORM.TEPRUORI.Caption = ATRIP$
     AMATEX = EDITFORM.TEPRUORI.Width * REDUFORMU
     EDITFORM.TEPRURED.Font = "Arial"
     If InStr(UCase$(EDITFORM.TEPRUORI.FontName), "COURIER") > 0 Then
        EDITFORM.TEPRURED.FontName = "Courier New"
     End If
     EDITFORM.TEPRURED.FontBold = False
     EDITFORM.TEPRURED.Caption = ATRIP$
     For KKKK = 24 To 1 Step -1
       EDITFORM.TEPRURED.FontSize = KKKK / 2
       If EDITFORM.TEPRURED.Width <= AMATEX Then GoTo 17
     Next KKKK
     '
17   EDITFORM.Picture2(PAGINACTIVA%).FontName = EDITFORM.TEPRURED.FontName '"Arial"
     EDITFORM.Picture2(PAGINACTIVA%).FontSize = KKKK / 2
     EDITFORM.Picture2(PAGINACTIVA%).FontBold = False
     ATRIP$ = RTrim$(ATRIP$)
     If Right$(ATRIP$, 1) = ";" Then
         ATRIP$ = Left$(ATRIP$, Len(ATRIP$) - 1)
         EDITFORM.Picture2(PAGINACTIVA%).Print ATRIP$;
        Else
         EDITFORM.Picture2(PAGINACTIVA%).Print ATRIP$
     End If
   Case 20
     EDITFORM.Picture2(PAGINACTIVA%).FontUnderline = XVALO(TRIM$(ATRIP$))
   Case 21
     Call INCLUDEIMAGE(ATRIP$, "SHOW")
   End Select
   On Error GoTo 0
   '
End Sub
'

Private Sub VScroll1_Scroll()
   Picture1(PAGINACTIVA%).Top = (-100) * VScroll1.Value
End Sub
