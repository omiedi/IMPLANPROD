VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form FORXMAIL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   840
      Top             =   105
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
      Left            =   105
      Top             =   105
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
End
Attribute VB_Name = "FORXMAIL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VieneForm As Boolean
Private Sub Form_Load()
    '
    CMD$ = TRIM$(Command$)
    MODENVI = CMD$
    If CMD$ = "" Then
       MAI_AUTO% = 0       ' no es automatico
       MAI_TEST% = 0       ' NO ES ENTORNO DE TEST
       MAI_REPT% = 0
       WINMAIL09.Show 1
       GoTo 99
    End If
    '
    If LUCOM$ = "" Then LUCOM$ = "C:\FLEXOFT\"
    '
    If Len(CMD$) < 12 Then
      MsgBox "FLXMAIL Mal Invocado - Falta Archivo de Transferencia", vbSystemModal
      GoTo 99
    End If
    '
    CLA1$ = Left$(CMD$, 6)
    CLA2$ = Mid$(CMD$, 7, 6)
    If CLA2$ <> CONTRACLA$(CLA1$) Then
      MsgBox "FLXMAIL Mal Invocado - Error Santo y Seña"
      GoTo 99
    End If
    '
    PPXX% = InStr(UCase$(CMD$), "/AUTO")           ' ENVIO AUTOMATICO SIN PRESENTAR PANTALLA
    If PPXX% > 0 Then
       MAI_AUTO% = 1
       CMD$ = Left$(CMD$, PPXX% - 1) + Mid$(CMD$, PPXX% + 5)
    End If
    '
    PPXX% = InStr(UCase$(CMD$), "/TEST")           ' TEST - REEMPLAZA DESTINO POR ebauer@flexoft.com.ar
    If PPXX% > 0 Then
       MAI_TEST% = 1
       CMD$ = Left$(CMD$, PPXX% - 1) + Mid$(CMD$, PPXX% + 5)
    End If
    '
    LUCOM$ = "C:\FLEXOFT\"
    If CMD$ <> "" Then
      LUCOM$ = Mid$(CMD$, 14)
    End If
    If LUCOM$ = "" Then LUCOM$ = "C:\FLEXOFT\"
    '
    ATRAMAI$ = LUCOM$ + CLA1$ + ".TMP"
    If EXISTE%(ATRAMAI$) < 1 Then
       MsgBox "No Existe " + ATRAMAI$
       GoTo 99
    End If
    '
    TTXX$ = LOADFILE$(ATRAMAI$)
    Call ENVIAMAIL(TTXX$)
    On Error Resume Next
    Kill Left$(ATRAMAI$, Len(ATRAMAI$) - 1) & "*"
    On Error GoTo 0
    '
99  Close: End
    '
End Sub

Sub ENVIAMAIL(TTXX$, Optional DATOS)
    '
    ' INICIO DE PROCESO DE ARMADO DEL MAIL
    '
    On Error Resume Next
    MAPISession1.DownLoadMail = False
    MAPISession1.NewSession = 1
    MAPISession1.SignOn
    MAPIMessages1.SessionID = MAPISession1.SessionID
    If Err Then
       On Error GoTo 0
       MsgBox "Imposible Enviar e-mail  !!!" + vbCrLf + vbCrLf + "Fallo en Proceso de Inicio" + vbCrLf + "del  Programa  de Correo.", vbSystemModal + vbCritical
       Exit Sub
    End If
    '
    ' LECTURA DE LOS DATOS RECIBIDOS PARA EL ARMADO DEL MAIL
    '
    DESTI$ = SUBTEX$(TTXX$, "DESTINO=", Chr$(124))
    DIREL$ = SUBTEX$(TTXX$, "DIRELEC=", Chr$(124))
    ASUNT$ = SUBTEX$(TTXX$, "ASUNTOM=", Chr$(124))
    TEXTO$ = SUBTEX$(TTXX$, "TEXTOMA=", Chr$(124))
    ADJUN$ = SUBTEX$(TTXX$, "ADJUNTO=", Chr$(124))
    SOLCO$ = SUBTEX$(TTXX$, "SOLCONF=", Chr$(124))
    '
    ' CUANDO VIENE DEL FORMULARIO WINMAIL09
    '
    If VieneForm = True Then
       DESTI$ = TRIM$(WINMAIL09.Text1)
       DIREL$ = TRIM$(WINMAIL09.Text3)
       ASUNT$ = TRIM$(WINMAIL09.Text2)
       TEXTO$ = TRIM$(WINMAIL09.Text5)
       ADJUN$ = TRIM$(WINMAIL09.Label5)
       VieneForm = False
    End If
    '
    If Left$(TEXTO$, 3) = "###" Then
       TXXT$ = LOADFILE$(Mid$(TEXTO$, 4))
       If TXXT$ <> "" Then TEXTO$ = TXXT$
    End If
    '
    ' VALIDA LA DIRECCION DE CORREO
    '
    If DIREL$ = "" Or InStr(DIREL$, "@") < 1 Then
      MsgBox "Dir-Elec Destino no Válida"
      Exit Sub
    End If
    '
    ' VERIFICA QUE HAYA UN ASUNTO
    '
    If ASUNT$ = "" Then
      MsgBox "Falta Asunto"
      Exit Sub
    End If
    '
    If DESTI$ = "" Then DESTI$ = DIREL$
    '
    ' ARMA DEL MAIL (DIRECCIONES, ASUNTO, CUERPO, ADJUNTOS)
    ' *****************************************************
    '
    MAPIMessages1.Compose
    MAPIMessages1.MsgIndex = -1
    '  HASTA EL 10/09/2015 ESTABA EL CODIGO DE MAS ABAJO
    '  MAPIMessages1.RecipIndex = 0
    '  MAPIMessages1.RecipType = 1
    '  MAPIMessages1.RecipDisplayName = DIREL$
    ' *************************************************
    
    ' DIRECCIONES DE CORREO PARA EL ENVIO
    '
    If InStr(1, DIREL$, ";") > 0 Then
        DIRELAUX$ = DIREL$
        IDDIREL% = 0
25      POSI% = InStr(1, DIRELAUX$, ";")
        If POSI% > 0 Then
           DIRMAIL$ = Mid$(DIRELAUX$, 1, POSI% - 1)
           MAPIMessages1.RecipIndex = IDDIREL%
           MAPIMessages1.RecipType = 1
           MAPIMessages1.RecipDisplayName = DIRMAIL$
           IDDIREL% = IDDIREL% + 1
           DIRELAUX$ = TRIM$(Mid$(DIRELAUX$, POSI% + 1))
        Else
           DIRMAIL$ = TRIM$(DIRELAUX$)
           MAPIMessages1.RecipIndex = IDDIREL%
           MAPIMessages1.RecipType = 1
           MAPIMessages1.RecipDisplayName = DIRMAIL$
           GoTo 27
        End If
        GoTo 25
    Else
        MAPIMessages1.RecipIndex = 0
        MAPIMessages1.RecipType = 1
        MAPIMessages1.RecipDisplayName = DIREL$
27  End If
    'MAPIMessages1.RecipDisplayName = DESTI$
    'MAPIMessages1.RecipAddress = DIREL$
    
    ' ASUNTO Y CUERPO DEL MAIL
    '
    MAPIMessages1.MsgSubject = ASUNT$
    MAPIMessages1.MsgNoteText = TEXTO$
    '
    ' SOLICITUD DE CONFIRMACION
    '
    If UCase$(SOLCO$) = "SI" Then
       MAPIMessages1.MsgReceiptRequested = True
    End If
    '
    ' ARCHIVOS ADJUNTOS
    '
    II& = 0
    While Len(ADJUN$) > 0
      PPXX% = InStr(ADJUN$, "&")
      If PPXX% < 1 Then PPXX% = 1 + Len(ADJUN$)
      ARAD$ = TRIM$(Left$(ADJUN$, PPXX% - 1))
      ADJUN$ = TRIM$(Mid$(ADJUN$, PPXX% + 1))
      ARAD1$ = ARAD$: If InStr(ARAD$, ".") < 1 Then ARAD1$ = ARAD$ + "."
      If ARAD$ <> "" Then
        If EXISTE%(ARAD1$) > 0 Then      ' PARA QUE ADMITA SIN EXTENSIÓN
          MAPIMessages1.AttachmentIndex = MAPIMessages1.AttachmentCount
          II& = II& + 1
          MAPIMessages1.AttachmentPosition = II& - 1
          MAPIMessages1.AttachmentPathName = ARAD$
        End If
      End If
    Wend
    '
    ' CIERRE DEL ARMADO Y ENVIO
    '
    On Error Resume Next
    '''MAPIMessages1.Action = 3   ' vbMessageSend
    '''MAPISession1.Action = 2    ' cierra sesion
    MAPIMessages1.Send False
    MAPISession1.SignOff
    On Error GoTo 0
    DoEvents
    '
End Sub

