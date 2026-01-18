VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form FORECMAI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FLEXOFT - Recepcion Automatica de Mails"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7890
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   7890
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List8 
      Height          =   1230
      Left            =   1680
      TabIndex        =   9
      Top             =   4935
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.FileListBox File8 
      Height          =   1065
      Left            =   3255
      TabIndex        =   8
      Top             =   4935
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.DirListBox Dir8 
      Height          =   1215
      Left            =   -105
      TabIndex        =   7
      Top             =   4935
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.FileListBox File1 
      Height          =   1260
      Left            =   4830
      TabIndex        =   6
      Top             =   4935
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Procesa Archivos Adjuntos"
      Height          =   435
      Left            =   2835
      TabIndex        =   5
      Top             =   4410
      Width           =   2430
   End
   Begin VB.Timer Timer2 
      Interval        =   60000
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Examina Bandeja de Entrada"
      Height          =   330
      Left            =   2835
      TabIndex        =   0
      Top             =   3360
      Width           =   2430
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   525
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   630
      Top             =   1785
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   327680
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   0
      Top             =   1785
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   327680
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Label7"
      Height          =   330
      Left            =   1680
      TabIndex        =   12
      Top             =   1785
      Width           =   4215
   End
   Begin VB.Label Label6 
      Caption         =   "Este timer se activa cuando hay un nuevo PDF y lo AppActiva  para mantenerlo en pantalla hasta que el usuario lo cierra."
      Height          =   435
      Left            =   525
      TabIndex        =   11
      Top             =   525
      Visible         =   0   'False
      Width           =   4635
   End
   Begin VB.Label Label5 
      Caption         =   "Este timer barre y revisa la bandeja de entrada cada 60 segundos"
      Height          =   435
      Left            =   525
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   2640
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7035
      TabIndex        =   4
      Top             =   525
      Width           =   750
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7035
      TabIndex        =   3
      Top             =   210
      Width           =   750
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Total de Mensajes:"
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
      Left            =   4830
      TabIndex        =   2
      Top             =   585
      Width           =   2115
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Leyendo Mensaje Nro:"
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
      Left            =   4830
      TabIndex        =   1
      Top             =   270
      Width           =   2115
   End
End
Attribute VB_Name = "FORECMAI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ESTABARRIENDO%
Dim ESTAVISIBLE%

Sub MUESTRAPDF(ARCHORI$)
   '
   Static ACROL$
   '
   If ACROL$ = "" Then
     ACROL$ = FILEFIND$("C:", "ACRORD32.EXE", 0)
     If ACROL$ = "" Then
       ACROL$ = "NONE"
     End If
   End If
   '
   If ACROL$ = "NONE" Then Exit Sub
   '
   RUTOR$ = ACROL$ + " "
   ANAME$ = RUTOR$ + ARCHORI$          ' + "PRUEBA.PDF"
   RETURNVALUE = Shell(ANAME$, 1)
   Timer1.Enabled = True
   '
End Sub

Private Sub Command1_Click()
   Call EXTRARJ("PEPE", OKX%)
End Sub

Private Sub Form_Load()
   If App.PrevInstance = True Then
     Close: End
     Exit Sub
   End If
   WindowState = 1
   Call BARREBANEN
End Sub

'
Private Sub Timer1_Timer()
    '
    On Error Resume Next
    TEXACT$ = "Adobe"
    AppActivate TEXACT$ + " Reader" ' Adobe Reader"
    SendKeys Chr$(27)
    If Err Then
      Timer1.Enabled = False
      Screen.MousePointer = 11
      TINI = Timer
      While Timer < TINI + 2: Wend
      Screen.MousePointer = 1
      ESTAVISIBLE% = 0
      Call BARREARCHREC
      Exit Sub
    End If
    ESTAVISIBLE% = 1
    On Error GoTo 0
    '
End Sub

Private Sub Command2_Click()
   '
   If ESTABARRIENDO% > 0 Then
     Exit Sub ' para que no entre dos veces simultaneamente
   End If
   '
   Label7 = "Barriendo Bandeja de Entrada ..."
   ESTABARRIENDO% = 1
   '
5  MAPISession1.DownLoadMail = True    ' False
   MAPISession1.LogonUI = False
   '
   On Error Resume Next
   MAPISession1.Action = 1
   If Err Then
     Label7 = ""
     ESTABARRIENDO% = 0
     Exit Sub
   End If
   On Error GoTo 0
   '
   MAPIMessages1.SessionID = MAPISession1.SessionID
   DoEvents
   '
   MAPIMessages1.Fetch
   '
99 On Error Resume Next
   MAPISession1.Action = 2  ' !!!
   On Error GoTo 0
   Label7 = ""
   ESTABARRIENDO% = 0
   '
End Sub

Sub BARREBANEN()
   '
   ' Este Modulo 'Barre' la bandeja de entrada de Outlook Express
   ' y si encuentra un mail enviado por FLEXOFT, validando clave y
   ' contraseña en MsgSubject pasa a 'BAJADJUN' para guardar los
   ' archivos adjuntos en 'C:\FLEXOFT\ADJUNREC'
   '
   ' Se activa por boton de comando o por 'Timer2' cada 1 minuto
   '
   If ESTABARRIENDO% > 0 Then
     Exit Sub ' para que no entre dos veces simultaneamente
   End If
   '
   Label7 = "Barriendo Bandeja de Entrada ..."
   ESTABARRIENDO% = 1
   '
   Dim TPAR(32)
   UMALEIDO$ = LOADFILE$("C:\FLEXOFT\UMALEIDO.TXT")
   UJ& = 0
   While InStr(UMALEIDO$, "|") > 0
     PPXX% = InStr(UMALEIDO$, "|")
     If PPXX% < 1 Then GoTo 5
     UJ& = UJ& + 1
     If UJ& > 32 Then GoTo 5
     TPAR(UJ&) = Trim$(Left$(UMALEIDO$, PPXX% - 1))
     UMALEIDO$ = Mid$(UMALEIDO$, PPXX% + 1)
   Wend
   '
5  MAPISession1.DownLoadMail = False
   MAPISession1.LogonUI = False
   '
   On Error Resume Next
   MAPISession1.Action = 1
   If Err Then
     Label7 = ""
     ESTABARRIENDO% = 0
     Exit Sub
   End If
   On Error GoTo 0
   '
   MAPIMessages1.SessionID = MAPISession1.SessionID
   DoEvents
   '
   MAPIMessages1.Fetch
   CAMENSA& = MAPIMessages1.MsgCount
   '
   UMENLE_INDI = Val(TPAR(1))
   If UMENLE_INDI < 1 Or UMENLE_INDI > CAMENSA& Then
     INI& = 1
     GoTo 7
   End If
   '
   MAPIMessages1.MsgIndex = UMENLE_INDI - 1
   If TPAR(2) = MAPIMessages1.MsgOrigAddress Then
     If TPAR(3) = MAPIMessages1.MsgDateReceived Then
       If TPAR(4) = MAPIMessages1.MsgSubject Then
         INI& = UMENLE_INDI + 1
       End If
     End If
   End If
   '
7  ARCHORI$ = ""
   Label4 = CAMENSA&
   DoEvents
   For U& = INI& To CAMENSA&
     Label3 = U&
     DoEvents
     On Error Resume Next
     MAPIMessages1.MsgIndex = U& - 1
     If Err Then
       On Error GoTo 0
       GoTo 99
     End If
     On Error GoTo 0
     '
     TTX1$ = MAPIMessages1.MsgOrigAddress
     ASUNTO$ = MAPIMessages1.MsgSubject
     '
     If Left$(Trim$(UCase$(ASUNTO$)), 7) = "FLEXOFT" Then
       PPXX% = InStr(ASUNTO$, "/")
       If PPXX% > 6 Then
         CLA1$ = Mid$(ASUNTO$, PPXX% - 6, 6)
         CLA2$ = Mid$(ASUNTO$, PPXX% + 1, 6)
         If CLA2$ = CONTRACLA$(CLA1$) Then    ' VALIDó ELECTRONICAMENTE EL MAIL
           '
           TPAR(2) = MAPIMessages1.MsgOrigAddress
           TPAR(3) = MAPIMessages1.MsgDateReceived
           TPAR(4) = MAPIMessages1.MsgSubject
           '
           TTXX$ = ""
           For KKI% = 2 To 4
             TTXX$ = TTXX$ & TPAR(KKI%) & "|"
           Next KKI%
           NARCHI$ = TTXX$
           '
           If EXISTE%("C:\FLEXOFT\MAILREC.IXX") > 0 Then
             N1% = FILEOPEN%("C:\FLEXOFT\MAILREC.IXX", 1, 128)
             While Not EOF(N1%)
               Line Input #N1%, NARCHJ$
               If NARCHJ$ = NARCHI$ Then
                 Close #N1%
                 GoTo 89   ' YA LO BAJARON
               End If
             Wend
             Close #N1%
           End If
           N1% = FILEOPEN%("C:\FLEXOFT\MAILREC.IXX", 3, 128)
           Print #N1%, NARCHI$
           Close #N1%
           'MsgBox NARCHI$
           '
           Call BAJADJUN
           '
         End If
       End If
     End If
     '
89   TPAR(1) = U&
     TPAR(2) = MAPIMessages1.MsgOrigAddress
     TPAR(3) = MAPIMessages1.MsgDateReceived
     TPAR(4) = MAPIMessages1.MsgSubject
     '
     TTXX$ = ""
     For KKI% = 1 To 4
       TTXX$ = TTXX$ & TPAR(KKI%) & "|"
     Next KKI%
     Call SAVEFILE("C:\FLEXOFT\UMALEIDO.TXT", TTXX$)
     '
     If ARCHORI$ <> "" Then
       MAPISession1.Action = 2
       Call MUESTRAPDF(ARCHORI$)
       Label7 = ""
       ESTABARRIENDO% = 0
       Exit Sub
     End If
     '
90 Next U&
   '
99 On Error Resume Next
   MAPISession1.Action = 2  ' !!!
   On Error GoTo 0
   Label7 = ""
   ESTABARRIENDO% = 0
   '
End Sub

Sub BAJADJUN()
   '
10 NADJULAR$ = MAPIMessages1.AttachmentName
   If NADJULAR$ <> "" Then
     NADJULAR$ = MAPIMessages1.AttachmentName
     
     NADJUCOR$ = NADJULAR$
     While InStr(NADJUCOR$, "\") > 0
       PPXX% = InStr(NADJUCOR$, "\")
       NADJUCOR$ = Mid$(NADJUCOR$, PPXX% + 1)
     Wend
     '
     If Len(NADJUCOR$) > 0 Then
       EXTEX$ = ""
       PPXX% = InStr(NADJUCOR$, ".")
       If PPXX% < 1 Then PPXX% = 1 + Len(NADJUCOR$)
       EXTEX$ = UCase$(Mid$(NADJUCOR$, PPXX% + 1))
       If InStr(EXTEX$, ".") > 0 Then
         GoTo 89    ' PARECE SER UN VIRUS
       End If
       '
       On Error Resume Next
       MkDir "C:\FLEXOFT\ADJUNREC"
       On Error GoTo 0
       '
       LAREXIS& = 0: AREXI$ = ""
       On Error Resume Next
       LAREXIS& = FileLen("C:\FLEXOFT\ADJUNREC\" + NADJUCOR$)
       On Error GoTo 0
       If LAREXIS& > 0 Then
         AREXI$ = LOADFILE$("C:\FLEXOFT\ADJUNREC\" + NADJUCOR$)
       End If
       ARREC$ = LOADFILE$(NADJULAR$)
       If ARREC$ <> AREXI$ Then
         'MsgBox "Recibo C:\FLEXOFT\ADJUNREC\" + NADJUCOR$
         FileCopy NADJULAR$, "C:\FLEXOFT\ADJUNREC\" + NADJUCOR$
       End If
       On Error GoTo 0
     End If
     '
89   On Error Resume Next
     MAPIMessages1.AttachmentIndex = MAPIMessages1.AttachmentIndex + 1
     If Err < 1 Then
       On Error GoTo 0
       GoTo 10
     End If
     On Error GoTo 0
   End If
   '
End Sub

Private Sub Command3_Click()
   '
   Call BARREARCHREC
   
End Sub

Sub BARREARCHREC()
   '
   Static UENTRA
   If Timer >= UENTRA Then
     If Timer < UENTRA + 3 Then
       Exit Sub
     End If
   End If
   If ESTAVISIBLE% > 0 Then
     Exit Sub
   End If
   '
   UENTRA = Timer
   '
   On Error Resume Next
   MkDir "C:\FLEXOFT"
   MkDir "C:\FLEXOFT\ADJUNREC"
   On Error GoTo 0
   '
   File1.Path = "C:\FLEXOFT\ADJUNREC"
   File1.Pattern = "*.*"
   File1.Refresh
   '
   CARCHIRE& = File1.ListCount
   Label4 = CARCHIRE&: DoEvents
   '
   For U& = 1 To CARCHIRE&
     SALEBARRE% = 0
     Label3 = U&: DoEvents
     NARCHICO$ = File1.List(U& - 1)
' If InStr(NARCHICO$, "ARJ") > 0 Then
' AAA = BBB
' End If
     If EXISTE%("C:\FLEXOFT\ADJUNREC\" + NARCHICO$) > 0 Then
       NARCHI$ = Format$(FileDateTime("C:\FLEXOFT\ADJUNREC\" + NARCHICO$), "DD-MM-YY HH:MM") + NARCHICO$
       '
       If EXISTE%("C:\FLEXOFT\ADJUNREC.IXX") > 0 Then
         N1% = FILEOPEN%("C:\FLEXOFT\ADJUNREC.IXX", 1, 128)
         While Not EOF(N1%)
           Line Input #N1%, NARCHJ$
           If NARCHJ$ = NARCHI$ Then
             Close #N1%
             GoTo 99
           End If
         Wend
         Close #N1%
       End If
       '
20     EXTEX$ = ""
       PPXX% = InStr(NARCHICO$, ".")
       If PPXX% < 1 Then PPXX% = 1 + Len(NADJUCOR$)
       EXTEX$ = UCase$(Mid$(NARCHICO$, PPXX% + 1))
       If InStr(EXTEX$, ".") > 0 Then
         NARCHICO$ = EXTEX$
         GoTo 20
       End If
       '
' MsgBox "Archivo Recibido de FLEXOFT: " + NARCHICO$
       '
       If EXTEX$ = "ARJ" Then
           Call EXTRARJ("C:\FLEXOFT\ADJUNREC\" + NARCHICO$, OKX%)
         ElseIf EXTEX$ = "ZIP" Then
         
         ElseIf EXTEX$ = "PDF" Then
           Call MUESTRAPDF("C:\FLEXOFT\ADJUNREC\" + NARCHICO$)
           SALEBARRE% = 1
         ElseIf EXTEX$ = "EXE" Then
           Call TRANOVE(NARCHICO$)
         ElseIf EXTEX$ = "PRF" Then
         
         ElseIf EXTEX$ = "FRM" Then
         
         ElseIf EXTEX$ = "RFS" Then
         
         Else
' MsgBox "Sin Instrucciones de Procesamiento para Archivo " + NARCHICO$
       End If
       '
       'AQUI LO MARCA COMO PROCESADO
       N1% = FILEOPEN%("C:\FLEXOFT\ADJUNREC.IXX", 3, 128)
       Print #N1%, NARCHI$
       Close #N1%
'MsgBox NARCHI$, 4096
       If SALEBARRE% = 1 Then Exit Sub
       '
     End If
     '
99 Next U&
   '
End Sub
'
Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   MAPISession1.Action = 2  ' !!!
   On Error GoTo 0
   Close: End
End Sub

'
Private Sub Timer2_Timer()
   '
   Call BARREBANEN
   Call BARREARCHREC
   '
End Sub

Function FILEFIND$(ULO$, ABUS$, MODO%)
   '
   SCMA = Screen.MousePointer
   '
   FFII$ = ""
   ESTEQ$ = ULO$
   Screen.MousePointer = 11
   '
   ABUSCA$ = ABUS$
   Dir8.Path = ESTEQ$ + "\"
   File8.Path = Dir8.Path
   File8.Pattern = ABUSCA$
   DIEXPLO$ = ESTEQ$ + "\"
   '
5  ' CARGA EN LIST1 LA LISTA DE LAS CARPETAS A EXPLORAR *************
   List8.Clear
   ILISTA& = 0
   List8.AddItem DIEXPLO$
   '
10 Dir8.Path = List8.List(ILISTA&)
   Dir8.Refresh
   DoEvents
   For KKI% = 1 To Dir8.ListCount
     TTXX$ = Dir8.List(KKI% - 1) + "\"
     List8.AddItem TTXX$
     List8.ListIndex = List8.ListCount - 1
     List8.Refresh
     DoEvents
   Next KKI%
   '
   If ILISTA& < List8.ListCount - 1 Then
     ILISTA& = ILISTA& + 1
     GoTo 10
   End If
   '
   For ILISTA& = 1 To List8.ListCount
     List8.ListIndex = ILISTA& - 1
     File8.Path = List8.List(ILISTA& - 1)
     DoEvents
     File8.Pattern = ABUSCA$
     '
     For KKI% = 1 To File8.ListCount
       FINAME$ = Trim$(UCase$(File8.List(KKI% - 1)))
       If FINAME$ <> "" Then
         ACOMPLE$ = List8.List(ILISTA& - 1) + FINAME$
         FILEFIND$ = ACOMPLE$
         Screen.MousePointer = SCMA
         Exit Function
       End If
29   Next KKI%
   Next ILISTA&
   Screen.MousePointer = SCMA
   '
End Function

Sub EXTRARJ(ARCHORI$, OKX%)
   '
   Static ACROL$
   '
   OKX% = 0
   '
   CRLF$ = Chr$(13) + Chr$(10)
   TTXX$ = "Se ha Recibido Correo de FLEXOFT, con" + CRLF$
   TTXX$ = TTXX$ + "un Adjunto en Formato 'ARJ'." + CRLF$
   TTXX$ = TTXX$ + "  " + CRLF$
   TTXX$ = TTXX$ + "¿ Acepta Extraer Archivos del Adjunto ?" + CRLF$
   '
   AB1% = MsgBox(TTXX$, 4145, "Correo FLEXOFT")
   If AB1% <> 1 Then GoTo 99
   '
   If ACROL$ = "" Then
     ACROL$ = FILEFIND$("C:", "ARJ.EXE", 0)
     If ACROL$ = "" Then
       ACROL$ = "NONE"
     End If
   End If
   '
   If ACROL$ = "NONE" Then
     TTXX$ = "Extracción Imposible. Falta Utilitario 'ARJ.EXE'" + CRLF$
     TTXX$ = TTXX$ + "en Sistema Instalado." + CRLF$
     TTXX$ = TTXX$ + "  " + CRLF$
     TTXX$ = TTXX$ + "Comuníquese con FLEXOFT para Resolver este" + CRLF$
     TTXX$ = TTXX$ + "Inconveniente." + CRLF$
     TTXX$ = TTXX$ + "  " + CRLF$
     TTXX$ = TTXX$ + "Telefax: (011)-4541-1515 / 4524-1284" + CRLF$
     '
     AB1% = MsgBox(TTXX$, 4144, "Falta Utilitario ARJ")
     Exit Sub
   End If
   '
   RUTOR$ = ACROL$ + " "
   CUDI$ = App.Path
   '
   LUDA$ = LUDAT$
   If Mid$(LUDA$, 2, 1) = ":" Then LUDA$ = Mid$(LUDA$, 3)
   If Left$(LUDA$, 2) = "\\" Then LUDA$ = Mid$(LUDA$, 2)
   '
   On Error Resume Next
20 CMD$ = ""
   CMD$ = CMD$ + "C:" + CRLF$
   CMD$ = CMD$ + "CHDIR \FLEXOFT\ADJUNREC " + CRLF$
   CMD$ = CMD$ + RUTOR$ + " E " + ARCHORI$ + " -O" + CRLF$
   If Mid$(CUDI$, 2, 1) = ":" Then
     CMD$ = CMD$ + Left$(CUDI$, 2) + CRLF$
     CUDI$ = Mid$(CUDI$, 3)
   End If
   '
   CMD$ = CMD$ + "CHDIR " + CUDI$ + CRLF$
   CMD$ = CMD$ + "EXIT " + CRLF$
   Call SAVEFILE("C:\FLEXOFT\DOSCOMM.BAT", CMD$)
   '
   MODOBA = vbMinimizedFocus
   SBX1 = Shell("C:\FLEXOFT\DOSCOMM.BAT", MODOBA)
   '
99 End Sub

Sub TRANOVE(ARCHORICOR$)
   '
   On Error Resume Next
   MkDir "\FLEXOFT\NOVEDAD"
   On Error GoTo 0
   '
   CUDI$ = App.Path
   On Error Resume Next
   FileCopy "C:\FLEXOFT\ADJUNREC\" + ARCHORICOR$, Left$(CUDI$, 2) + "\FLEXOFT\NOVEDAD\" + ARCHORICOR$
   If Err Then
     On Error GoTo 0
     MsgBox "Imposible Copiar " + ARCHORICOR$ + " a \FLEXOFT\NOVEDAD"
   End If
   On Error GoTo 0
   '
End Sub

