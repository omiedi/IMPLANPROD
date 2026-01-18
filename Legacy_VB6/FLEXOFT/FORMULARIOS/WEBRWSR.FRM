VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form WEBROWSR 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FLEXOFT - Revision automatica de la WEB"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   6240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Height          =   375
      Left            =   5500
      Picture         =   "WEBRWSR.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Bajada Directa de Archivos"
      Top             =   360
      Width           =   350
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   1440
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   840
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   400
      Left            =   0
      ScaleHeight     =   405
      ScaleWidth      =   435
      TabIndex        =   11
      Top             =   0
      Width           =   430
      Begin VB.Image Image3 
         Height          =   240
         Left            =   105
         Picture         =   "WEBRWSR.frx":030A
         ToolTipText     =   "Doble-Click para Des-Habilitar Web Services en esta Estación de Trabajo"
         Top             =   105
         Width           =   240
      End
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   5825
      Picture         =   "WEBRWSR.frx":114C
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Refresca la Informacion"
      Top             =   360
      Width           =   350
   End
   Begin VB.ListBox List3 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2460
      IntegralHeight  =   0   'False
      ItemData        =   "WEBRWSR.frx":1296
      Left            =   105
      List            =   "WEBRWSR.frx":1298
      TabIndex        =   5
      Top             =   1575
      Width           =   6000
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   3990
      Top             =   315
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "999"
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
      Left            =   5800
      TabIndex        =   9
      ToolTipText     =   "Doble Click para Restablecer Web Services"
      Top             =   105
      Width           =   400
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
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
      Left            =   105
      TabIndex        =   8
      Top             =   525
      Width           =   6000
   End
   Begin VB.Label Label5 
      Caption         =   "Este timer barre y revisa la WEB CADA 10 MINUTOS"
      Height          =   435
      Left            =   4410
      TabIndex        =   0
      Top             =   315
      Visible         =   0   'False
      Width           =   2640
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   6195
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label3 
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
      Height          =   300
      Left            =   105
      TabIndex        =   3
      Top             =   315
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   $"WEBRWSR.frx":129A
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
      Left            =   105
      TabIndex        =   7
      Top             =   945
      Width           =   6000
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   $"WEBRWSR.frx":1369
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
      Left            =   105
      TabIndex        =   6
      Top             =   4095
      Width           =   6000
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
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
      Left            =   105
      TabIndex        =   4
      Top             =   315
      Width           =   6000
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   105
      TabIndex        =   2
      Top             =   105
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      TabIndex        =   1
      Top             =   105
      Width           =   6000
   End
End
Attribute VB_Name = "WEBROWSR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Command1_Click()
   Command1.Enabled = False
   Timer2.Enabled = False
   Call VERIACWEB(1)
   Call SHOWLINEWS
   Timer2.Enabled = True
   Command1.Enabled = True
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   Timer2.Enabled = False
   Call BAJADIREC
   Timer2.Enabled = True
   Command2.Enabled = True
End Sub

Private Sub Form_Load()
   '
   On Error Resume Next
   FIDATEX = FileDateTime(App.Path + "\" + App.EXEName + ".EXE")
   On Error GoTo 0
   '
   ' Verifica si ya se está ejecutando pro PREVINSTANCE
   If App.PrevInstance = True Then
      CMD$ = Command$
      If CMD$ <> "REFRESH" Then
         Close: End: Exit Sub
      End If
   End If
   '
   ' Vuelve a verificar por archivo abierto en C:\FLEXOFT
   NK% = FreeFile
   On Error Resume Next
   Open "C:\FLEXOFT\FLX_BRWS.EXK" For Output Lock Read Write As #NK%
   Print #NK%, Time$
   If Err Then
      ERNU% = Err.Number     ' Si es 71 es que no existe C:\FLEXOFT; entonces sigue
      On Error GoTo 0
      If ERNU% = 52 Or ERNU% = 55 Then    ' Ya está abierto y en ejecucion
         Close: End: Exit Sub
      End If
   End If
   On Error GoTo 0
   '
   On Error Resume Next
   Caption = "WEB Services FLEXOFT - " + Format(FileDateTime(App.Path + "\FLX_BRWS.EXE"), "dd/mm/yy  hh:mm")
   On Error GoTo 0
   WindowState = 1           ' ejecuta minimizado
   '
   Call LEEDATSYS
   Call CREABACONLEC
   '
   If IDEPLI$ = "41D8-F7FA" Then
      Timer2.Interval = 50  ' FUERZA BAJA C/30 SEGS EN NOTEBOOK
   End If
   '
   Label1 = NEPLI$ + " (" + IDEPLI$ + ")"
      TTXX$ = Trim$(Str$(USERNU%)) + " - " + Trim$(USERNA$) + " (" + Trim$(USERID$) + ")"
   Label4 = TTXX$
   Label8 = CARPEDOWN$
   '
   Call TACHAWEB
   '
   Timer2.Enabled = False
   Call VERIACWEB(0)
   Call SHOWLINEWS
   Timer2.Enabled = True
   '
End Sub

Sub TACHAWEB()
   '
   Picture3.Visible = False
   WEBSE$ = "DISABLED"
   If PUEDEBAJARDELAWEB% = 1 Then
      Picture3.Visible = True
   End If
   '
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Cancel = False
   If UnloadMode = 0 Then     ' desde el botón x
      WindowState = 1         ' minimiza
      Cancel = True           ' cancela la descarga
   End If
End Sub
'
Private Sub Form_Unload(Cancel As Integer)
   Close
   End
End Sub

Private Sub Image3_DblClick()
   
   WEBSE$ = "DISABLED"
   AKX$ = LOADFILE$("C:\FLEXOFT\WEB_SER.TXT")
   If AKX$ <> "" Then
      Call SAVEFILE("C:\FLEXOFT\WEB_SER.TXT", WEBSE$)
      Call TACHAWEB
   End If
   '
End Sub

Private Sub Label1_DblClick()
   Unload Me
End Sub

Private Sub Label9_DblClick()
   '
   WEBSE$ = "ENABLED"
   AKX$ = LOADFILE$("C:\FLEXOFT\FLEXRAM.DRV")
   If AKX$ <> "" Then
        Call SAVEFILE("C:\FLEXOFT\WEB_SER.TXT", WEBSE$)
        Call TACHAWEB
      Else
        MsgBox "No es  Posible  Establecer WEB Services  en esta  Estación de Trabajo." + vbCrLf + vbCrLf + "Causa Probable: Acceso por Terminal Services - No Existe 'C:\FLEXOFT'"
   End If
   '
End Sub

Private Sub List3_DblClick()
    Timer2.Enabled = False
    ARNAME$ = UCase$(Trim$(Left$(List3.Text, 32)))
    ARFULLNAME$ = AJUSTI$(List3.Text, 48)
    If InStr(ARNAME$, ".PDF") > 0 Then
        Call MUESTRAPDF(ARNAME$, ARFULLNAME$)
      ElseIf InStr(ARNAME$, ".DOC") > 0 Then
        Call MUESTRADOC(ARNAME$, ARFULLNAME$)
      ElseIf InStr(ARNAME$, ".XLS") > 0 Then
        Call MUESTRADOC(ARNAME$, ARFULLNAME$)
      Else
        Call INSTALACON(ARFULLNAME$)
    End If
    Timer2.Enabled = True
End Sub

Sub MUESTRAPDF(ARNAME$, ARFULLNAME$)
    '
    IDME$ = Trim$(Mid$(List3.Text, 65, 4))
    '
    On Error Resume Next
    SHOWPDF.List1.Clear
    SHOWPDF.Label1.Caption = ARFULLNAME$ + AJUSTI$(USERID$, 16)
    SHOWPDF.Pdf1.LOADFILE$ (CARPEDOWN$ + ARNAME$)
    If Err Then
         On Error GoTo 0
         PROCFILE.List1.Clear
         PROCFILE.Label1.Caption = ARFULLNAME$ + AJUSTI$(USERID$, 16)
         PROCFILE.Label4 = ARNAME$
         PROCFILE.Label5 = IDME$
         If IDME$ <> "" Then
            For JJ1& = 1 To List3.ListCount
               IDME1$ = Trim$(Mid$(List3.List(JJ1& - 1), 65, 4))
               If IDME1$ = IDME$ Then
                  If AJUSTI$(List3.List(JJ1& - 1), 48) <> ARFULLNAME$ Then
                     PROCFILE.List1.AddItem AJUSTI$(List3.List(JJ1& - 1), 48)
                  End If
               End If
            Next JJ1&
         End If
         Call MUESTRADOC(ARNAME$, ARFULLNAME$)
      Else
         On Error GoTo 0
         SHOWPDF.Label4 = ARNAME$
         SHOWPDF.Label5 = IDME$
         If IDME$ <> "" Then
            For JJ1& = 1 To List3.ListCount
               IDME1$ = Trim$(Mid$(List3.List(JJ1& - 1), 65, 4))
               If IDME1$ = IDME$ Then
                  If AJUSTI$(List3.List(JJ1& - 1), 48) <> ARFULLNAME$ Then
                     SHOWPDF.List1.AddItem AJUSTI$(List3.List(JJ1& - 1), 48)
                  End If
               End If
            Next JJ1&
         End If
         SHOWPDF.Show 1
    End If
    '
End Sub
'
Sub MUESTRADOC(ARNAME$, ARFULLNAME$)
    '

    PROCFILE.Label1.Caption = ARFULLNAME$ + AJUSTI$(USERID$, 16)
    ' Ejecutar un acceso directo
    Call ShellExecute(Me.hwnd, "Open", CARPEDOWN$ + ARNAME$, "", "", 1)
 
    PROCFILE.Label2 = Trim$(Left$(List3.Text, 48))
    PROCFILE.Show
    '
End Sub
'
Sub INSTALACON(ARFULLNAME$)
    '
    NARCHX$ = UCase$(Trim$(Left$(ARFULLNAME$, 16)))
    If NARCHX$ = "FLXMAIL.EXE" Or NARCHX$ = "FLOGON.EXE" Or NARCHX$ = "FLX_BRWS.EXE" Or InStr(NARCHX$, ".REL") > 0 Then GoTo 40
    '
    LFI$ = LOADFILE$("C:\FLEXOFT\FLEXRAM.DRV")
    USERNU% = CVI(Mid$(LFI$, 9, 2)) Mod 100
    If USERNU% < 1 Then
       MsgBox "Operacion Cancelada - Imposible Identificar Usuario Activo.", vbCritical, "FLEXOFT WEB Services"
       Exit Sub
    End If
    LUDAT$ = Trim$(Mid$(LFI$, 99, 46))
    '
    MDY$ = "BAJACTU"
    QWR% = USERNU% Mod 100
    If QWR% <> 1 Then
       FAXX$ = LOADFILE$(LUDAT$ + "FLEXACC.DAT")
       For KU& = 1 To Len(FAXX$) Step 48
          XTY$ = Mid$(FAXX$, KU&, 48)
          If CVI(Left$(XTY$, 2)) = QWR% Then
            If Trim$(UCase$(Mid$(XTY$, 3, 8))) = MDY$ Then
              If Mid$(XTY$, 11, 1) = "S" Then DRXXX% = 1
              If Mid$(XTY$, 12, 1) = "S" Then DRXXX% = 2
              GoTo 40
            End If
          End If
       Next KU&
       '
       FAXX$ = LOADFILE$(LUDAT$ + "INDIAPLI.DAT")
       For KU& = 1 To Len(FAXX$) Step 48
          XTY$ = Mid$(FAXX$, KU&, 48)
          If Trim$(Left$(XTY$, 12)) = MDY$ Then GoTo 30
       Next KU&
       FAXX$ = FAXX$ + AJUSTI$(MDY$, 12) + AJUSTI$("Instalacion de Novedades", 36)
       Call SAVEFILE(LUDAT$ + "INDIAPLI.DAT", FAXX$)
       '
30     MsgBox "Usuario no habilitado para instalacion de novedades bajadas de la WEB. Solicite la habilitacion correspondiente al supervisor de la red.", vbCritical, "FLEXOFT WEB Services"
       Exit Sub
       '
    End If
    '
40  DIREPRO$ = App.Path
    If EXISTE%("F:\FLEXOFT\NOVEDAD\.") > 0 Then
         DIRENEW$ = "F:\FLEXOFT\NOVEDAD"
       Else
         DIRENEW$ = PARENTFOLDER$(DIREPRO$) + "\NOVEDAD"
    End If
    '
    ' CREA LA CARPETA \NOVEDAD
    On Error Resume Next
    MkDir DIRENEW$
    On Error GoTo 0
    '
    ARNAME$ = UCase$(Trim$(Left$(ARFULLNAME$, 32)))
    ARNAMDEST$ = ARNAME$: If InStr(ARNAME$, ".REL") > 0 Then ARNAMDEST$ = "MSTDRV04.EXE"
    NEWFILE$ = LOADFILE$(CARPEDOWN$ + ARNAME$)
    OLDFILE$ = LOADFILE$(DIRENEW$ + "\" + ARNAMDEST$)
    If NEWFILE$ <> "" Or InStr(ARNAME$, ".REL") > 0 Then
       If NEWFILE <> OLDFILE$ Then
          Call SAVEFILE(DIRENEW$ + "\" + ARNAMDEST$, NEWFILE$)
          If InStr(ARNAME$, ".REL") > 0 Then
             On Error Resume Next
             Call SAVEFILE(App.Path + "\" + ARNAMDEST$, NEWFILE$)
             On Error GoTo 0
          End If
       End If
       Call CONFIRMALEC(ARFULLNAME$)
       '
       If NARCHX$ <> "FLOGON.EXE" Then
         If NARCHX$ <> "FLX_BRWS.EXE" Then
           If NARCHX$ <> "FLXMAIL.EXE" Then
             MsgBox "Archivo '" + ARNAME$ + "' colocado en carpeta '" + DIRENEW$ + "'." + vbCrLf + vbCrLf + "Re-Inicie ahora la aplicación para auto-instalar.", vbCritical, "FLEXOFT WEB Services"
             GoTo 55
           End If
         End If
       End If
       '
       ' LOS 3 ANTERIORES SE INSTALAN DIRECTAMENTE EN APP.PATH
' MsgBox "*** Instalado " + NARCHX$ + " EN " + App.Path
       On Error Resume Next
       Call FileCopy(DIRENEW$ + "\" + ARNAMDEST$, App.Path + "\" + ARNAMDEST$)
       On Error GoTo 0
       '
55     PPXX% = InStr(NARCHX$, "."): If PPXX% < 2 Then PPXX% = 1 + Len(NARCHX$)
       If UCase$(Left$(NARCHX$, PPXX% - 1)) = UCase$(App.EXEName) Then
          MsgBox "Actualizacion de '" + NARCHX$ + "' Bajada de la WEB. Oportunamente Re-Inicie FLEXOFT para Instalar Novedad. Muchas Gracias.", vbSystemModal + vbExclamation
       End If
    End If
    '
End Sub

Sub CONFIRMALEC(ARFULLNAME$)
   '
   ARNA$ = Trim$(Left$(ARFULLNAME$, 32))
   FIDATI$ = Trim$(Mid$(ARFULLNAME$, 33, 16))
   '
   If ARNA$ = "" Or FIDATI$ = "" Then
      MsgBox "No fue Posible Enviar por Mail la Confirmación de Lectura / Instalacion. Faltan Datos del Archivo Instalado / Leído." + vbCrLf$ + "Se Recomienda Consultar Telefónicamente al (54-11) 4524-1284 / 4522-2048. Muchas Gracias.", vbSystemModal
      Exit Sub
   End If
   '
   SQLX$ = "SELECT * FROM CONTRO_DOWN"
   SQLX$ = SQLX$ + " WHERE File_Name = '" & ARNA$ & "' "
   SQLX$ = SQLX$ + " AND File_DateTime = '" & FIDATI$ & "'"
   SQLX$ = SQLX$ + " AND (FeHor_Lect = '' OR ISNULL(FeHor_Lect) = TRUE) "
   SQLX$ = SQLX$ + " AND (User_Dest = '" & USERID$ & "' OR User_Dest = '')"
   '
   Call CREABACONLEC
   Set mirs = BADACON.OpenRecordset(SQLX$, dbOpenDynaset)
   On Error Resume Next
   mirs.MoveFirst
   If Err Then
      On Error GoTo 0
      mirs.Close
      BADACON.Close
      GoTo 99                 ' no hay nada que bajar
   End If
   On Error GoTo 0
   '
   With mirs
      Do While Not .EOF
         .Edit
         AHORA$ = Format$(Now, "dd/mm/yy  hh:mm")
         !FeHor_Lect = AHORA$
         !Usu_Lect = USERNA$
         .Update
      .MoveNext
      Loop
   End With
   On Error Resume Next
   mirs.Close
   BADACON.Close
   On Error GoTo 0
   '
   ' AQUI ENVIAR UN MAIL CON LA CONFIRMACION DE LECTURA A FLEXOFT
   TEXTOMAI$ = "Confirmacion de Lectura / Instalacion:" + vbCrLf
   TEXTOMAI$ = TEXTOMAI$ + vbCrLf
   TEXTOMAI$ = TEXTOMAI$ + "Documento: " + Trim$(ARNA$) + vbCrLf
   TEXTOMAI$ = TEXTOMAI$ + "Empresa: " + NEPLI$ + vbCrLf
   TEXTOMAI$ = TEXTOMAI$ + "LAN-ID: " + IDEPLI$ + vbCrLf
   TEXTOMAI$ = TEXTOMAI$ + "Usuario: " + USERNA$ + vbCrLf
   TEXTOMAI$ = TEXTOMAI$ + "USER-ID: " + USERID$ + vbCrLf
   TEXTOMAI$ = TEXTOMAI$ + "Fecha: " + Trim$(FIDATI$) + vbCrLf
   TEXTOMAI$ = TEXTOMAI$ + "" + vbCrLf
   '
   ' MAIL DE CONFIRMACION DE LECTURA
   On Error GoTo 50
   MAPISession1.DownLoadMail = False
   MAPISession1.NewSession = True
   MAPISession1.SignOn
   MAPIMessages1.SessionID = MAPISession1.SessionID
   '
   MAPIMessages1.Compose
   MAPIMessages1.RecipDisplayName = "asistec@flexoft.com.ar"
   MAPIMessages1.RecipAddress = "asistec@flexoft.com.ar"
   MAPIMessages1.MsgSubject = "Flexoft WEB Services"
   MAPIMessages1.MsgNoteText = TEXTOMAI$
   '
   MAPIMessages1.Send False
   MAPISession1.SignOff
   On Error GoTo 0
   DoEvents
   GoTo 60
   '
50 Resume 51
   ' error de envio de mail
51 On Error GoTo 0
   MsgBox "No fue Posible Enviar por Mail la Confirmación de Lectura / Instalacion Correspondiente a '" + ARNA$ + " - " + FIDATI$ + "'. Se Recomienda Consultar Telefónicamente al (54-11) 4524-1284 / 4522-2048. Muchas Gracias.", vbSystemModal
   '
   '
60 ' ELIMINA DE LA LISTA DE SELECCION
   ARFU$ = AJUSTI$(ARFULLNAME$, 48)
65 For KKIL& = 1 To WEBROWSR.List3.ListCount
      If Left$(WEBROWSR.List3.List(KKIL& - 1), 48) = ARFU$ Then
         WEBROWSR.List3.RemoveItem KKIL& - 1
         WEBROWSR.Refresh
         GoTo 65
      End If
   Next KKIL&
      
99 End Sub
'
Private Sub Timer2_Timer()
   Static SEGUNDOS%, TINI
   '
   If TINI = 0 Then TINI = Timer
   SEGUNDOS% = (SEGUNDOS% + 1) Mod 600   ' BARRE CADA 10 MINUTOS
   Label9.Caption = Int(600 - SEGUNDOS%)
   If PROCFILE.Visible = True Then
        WindowState = 1
   End If
   If SEGUNDOS% > 0 Then Exit Sub
   '
   Timer2.Enabled = False
   Call VERIACWEB(0)
   Call SHOWLINEWS
   If List3.ListCount > 0 Then
      WindowState = 0
   End If
   Timer2.Enabled = True
End Sub

Sub SHOWLINEWS()
   '
   SQLX$ = "SELECT * FROM CONTRO_DOWN"
   SQLX$ = SQLX$ + " WHERE FeHor_Baja <> '' "
   SQLX$ = SQLX$ + " AND (FeHor_Lect = '' OR ISNULL(FeHor_Lect) = TRUE) "
   SQLX$ = SQLX$ + " AND (User_Dest = '" & USERID$ & "' OR User_Dest = '') "
   '
   Call CREABACONLEC
   Set mirs = BADACON.OpenRecordset(SQLX$, dbOpenSnapshot)
   On Error Resume Next
   mirs.MoveLast
   If Err Then
      mirs.Close
      BADACON.Close
      On Error GoTo 0
      GoTo 99                 ' no hay nada que bajar
   End If
   On Error GoTo 0
   '
   TTXX$ = "": J3& = 0: List3.Clear
   With mirs
      .MoveFirst
      Do While Not .EOF
        IDME$ = Trim$(!Num_Pack)
        ARNA$ = Trim$(UCase$(!File_Name))
        FIDATI$ = Trim$(!File_DateTime)
        FDTX$ = Mid$(FIDATI$, 7, 2) + Mid$(FIDATI$, 4, 2) + Mid$(FIDATI$, 1, 2) + "." + Trim$(Mid$(FIDATI$, 9))
        FDTX1$ = REPLACAR$(FDTX$, ":", "")
        FACTE# = Val(FDTX$) + (Val(FDTX1$)) / 10000
        '
        TTZZ$ = AJUSTI$(AJUSTI$(ARNA$, 32) + FIDATI$, 48)
        TTZZ1$ = AJUSTI$(TTZZ$ + USERID$, 64)
        TTZZ2$ = Trim$(Left$(TTZZ$, 32))
        '
        ' si es FLOGON.EXE o FLX_BRWS.EXE lo copia directamente A NOVEDAD
        If UCase$(TTZZ2$) = "FLXMAIL.EXE" Or UCase$(TTZZ2$) = "FLOGON.EXE" Or UCase$(TTZZ2$) = "FLX_BRWS.EXE" Then
           Call INSTALACON(TTZZ$)
           GoTo 19
        End If
        '
        ' esto se agrego el 20/05/2009 para que solo muestre los existentes
        ARCHJ$ = Trim$(Left$(ARNA$, 32))
        If EXISTE%(CARPEDOWN$ + ARCHJ$) > 0 Then
16         For KKI% = 1 To List3.ListCount
              If Trim$(UCase$(Left$(List3.List(KKI% - 1), 32))) = ARCHJ$ Then
                FIDATJ$ = Trim$(Mid$(List3.List(KKI% - 1), 33, 16))
                FDTY$ = Mid$(FIDATJ$, 7, 2) + Mid$(FIDATJ$, 4, 2) + Mid$(FIDATJ$, 1, 2) + "." + Trim$(Mid$(FIDATJ$, 9))
                FDTY1$ = REPLACAR$(FDTY$, ":", "")
                FANTE# = Val(FDTY$) + (Val(FDTY1$)) / 10000
                   If FANTE# > FACTE# Then GoTo 19
                   List3.RemoveItem (KKI% - 1)
                   GoTo 16
              End If
           Next KKI%
           List3.AddItem TTZZ$ + Space$(16) + IDME$
        End If
        '
19    .MoveNext
      Loop
   End With
   '
   If List3.ListCount > 0 Then
      Me.WindowState = 0
      Me.Visible = True
      Me.ZOrder 0
      AppActivate Caption
   End If
   '
99 On Error Resume Next
   mirs.Close
   Set mirs = Nothing
   BADACON.Close
   On Error GoTo 0
   '
End Sub
'
