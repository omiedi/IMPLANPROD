VERSION 5.00
Begin VB.Form FXMAIWSK09 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FLEXOFT - Correo Electronico"
   ClientHeight    =   810
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   3510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   810
   ScaleWidth      =   3510
   StartUpPosition =   3  'Windows Default
   Begin VB.Label MODENVI 
      Height          =   255
      Left            =   2760
      TabIndex        =   1
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Conectando ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   2895
   End
End
Attribute VB_Name = "FXMAIWSK09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LUCOM$
Dim MAI_AUTO%
Dim MAI_REPT%

Private Sub Form_Load()
    '
    CMD$ = TRIM$(Command$)
    MODENVI = CMD$
    If CMD$ = "" Then
       MAI_AUTO% = 0       ' no es automatico
       MAI_TEST% = 0       ' NO ES ENTORNO DE TEST
       MAI_REPT% = 0
       Call ENVIAMAIL(TTXX$)
       GoTo 99
    End If
    '
    If LUCOM$ = "" Then LUCOM$ = "C:\FLEXOFT\"
    '
    If Len(CMD$) < 12 Then
      MsgBox "FLXMAIL Mal Invocado - Falta Archivo de Transferencia", vbSystemModal
      Close: End
    End If
    '
    CLA1$ = Left$(CMD$, 6)
    CLA2$ = Mid$(CMD$, 7, 6)
    If CLA2$ <> CONTRACLA$(CLA1$) Then
      MsgBox "FLXMAIL Mal Invocado - Error Santo y Seña"
      Close: End
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
    LUCOMAI$ = LUCOM$
    If TRIM$(Mid$(CMD$, 14)) <> "" Then
      LUCOMAI$ = Mid$(CMD$, 14)
    End If
    '
10  ATRAMAI$ = LUCOMAI$ + CLA1$ + ".TMP"
    If EXISTE%(ATRAMAI$) > 0 Then
         TTXX$ = LOADFILE$(ATRAMAI$)
         Call ENVIAMAIL(TTXX$)
         If MAI_SENDOK > 0 Or FXWSKSEND09.Result = "OK" Then     ' SI NO PUDO ENVIARLO LO CONSERVA
               If MAI_REPT% > 0 Then
                     TTXX$ = LOADFILE$(ATRAMAI$)
                     Call SAVEFILE(ATRAMAI$, TTXX$)    ' PARA QUE PONGA FECHA Y HORA DE ENVIO
                   Else
                     On Error Resume Next
                     Kill LUCOMAI$ + CLA1$ + ".*"
                     On Error GoTo 0
               End If
               If MAI_AUTO% < 1 Then
                  MsgBox "Envío Exitoso"
               End If
               GoTo 99
            Else
               If MAI_AUTO% < 1 Then
                  OPX% = MsgBox("No pudo Enviarse el Mail Correspondiente a " + ATRAMAI$, vbAbortRetryIgnore)
                  If OPX% = 4 Then GoTo 10
               End If
         End If
       Else
         MsgBox "No Existe " + ATRAMAI$
    End If
    '
99  Close: End
    '
End Sub

Sub ENVIAMAIL(TTXX$)
  '
  DESTI$ = SUBTEX$(TTXX$, "DESTINO=", Chr$(124))
    FXWSKSEND09.txtToName.Text = DESTI$
  DIREL$ = SUBTEX$(TTXX$, "DIRELEC=", Chr$(124))
    If MAI_TEST > 0 Then DIREL$ = "asistec@flexoft.com.ar"
    FXWSKSEND09.txtTo.Text = DIREL$
  DESTICOP$ = SUBTEX$(TTXX$, "DESTICC=", Chr$(124))
    FXWSKSEND09.txtCcName.Text = DESTICOP$
  DIRELCOP$ = SUBTEX$(TTXX$, "DIRELCC=", Chr$(124))
    FXWSKSEND09.txtCc.Text = DIRELCOP$
  ASUNT$ = SUBTEX$(TTXX$, "ASUNTOM=", Chr$(124))
    FXWSKSEND09.txtSubject.Text = ASUNT$
  TEXTO$ = SUBTEX$(TTXX$, "TEXTOMA=", Chr$(124))
    If Left$(TEXTO$, 3) = "###" Then                ' ES LA REFERENCIA A UN ARCHIVO CON RUTA
       ATEX$ = TRIM$(Mid$(TEXTO$, 4))               ' COMPLETA
       If EXISTE%(ATEX$) > 0 Then
          TEXTO$ = LOADFILE$(ATEX$)
       End If
    End If
    FXWSKSEND09.txtMsg = TEXTO$
  FORTX$ = SUBTEX$(TTXX$, "FORTEXT=", Chr$(124))
    If FORTX$ = "HTML" Then FXWSKSEND09.ckHtml.Value = 1
  ADJUN$ = SUBTEX$(TTXX$, "ADJUNTO=", Chr$(124))
10    PPXX% = InStr(ADJUN$, "&")
      If PPXX% > 0 Then ADJUN$ = Left$(ADJUN$, PPXX% - 1) + ";" + Mid$(ADJUN$, PPXX% + 1): GoTo 10
      PPXX% = InStr(ADJUN$, " ;")
      If PPXX% > 0 Then ADJUN$ = Left$(ADJUN$, PPXX% - 1) + ";" + Mid$(ADJUN$, PPXX% + 2): GoTo 10
      PPXX% = InStr(ADJUN$, "; ")
      If PPXX% > 0 Then ADJUN$ = Left$(ADJUN$, PPXX% - 1) + ";" + Mid$(ADJUN$, PPXX% + 2): GoTo 10
    FXWSKSEND09.txtAttach.Text = ADJUN$
  SOLCO$ = SUBTEX$(TTXX$, "SOLCONF=", Chr$(124))
    If SOLCO$ = "SI" Then FXWSKSEND09.ckReceipt.Value = 1
  REPIT$ = TRIM$(SUBTEX$(TTXX$, "REPEAT=", Chr$(124)))
    If REPIT$ <> "" Then MAI_REPT% = 1
  FXWSKSEND09.ckLogin.Value = 1
  '
  'COMPLETA DATOS DEL MAIL
  '
  TTYY$ = LOADFILE(LUCOM$ + "\STPFMAIL.CFG")
  SSMTP$ = SUBTEX$(TTXX$, "SMTPSERVER=", Chr$(124))                 ' DEL ARCHIVO TRANSFERENCIA
    If SSMTP$ = "" Then SSMTP$ = TRIM$(Mid$(TTYY$, 1, 128))         ' DE LA CONFIGURACION BASICA
    If SSMTP$ = "" Then SSMTP$ = "SMTP.MAIL.YAHOO.COM.AR"           ' POR DEFAULT
    FXWSKSEND09.txtServer = SSMTP$
  SNAME$ = SUBTEX$(TTXX$, "SENDERNAME=", Chr$(124))                 ' DEL ARCHIVO TRANSFERENCIA
    If SNAME$ = "" Then SNAME$ = TRIM$(Mid$(TTYY$, 129, 128))       ' DE LA CONFIGURACION BASICA
    If SNAME$ = "" Then SNAME$ = "FLEXOFT Soluciones"               ' POR DEFAULT
    FXWSKSEND09.txtFromName = SNAME$
  SEMAI$ = SUBTEX$(TTXX$, "SENDERMAIL=", Chr$(124))                ' DEL ARCHIVO TRANSFERENCIA
    If SEMAI$ = "" Then SEMAI$ = TRIM$(Mid$(TTYY$, 257, 128))       ' DE LA CONFIGURACION BASICA
'    If SEMAI$ = "" Then SEMAI$ = "flexoftsoluciones@yahoo.com.ar"   ' POR DEFAULT
    FXWSKSEND09.txtFrom = SEMAI$
  UNAME$ = SUBTEX$(TTXX$, "USERNAME=", Chr$(124))                   ' DEL ARCHIVO TRANSFERENCIA
    If UNAME$ = "" Then UNAME$ = TRIM$(Mid$(TTYY$, 385, 128))       ' DE LA CONFIGURACION BASICA
    If UNAME$ = "" Then UNAME$ = "flexoftsoluciones@yahoo.com.ar"   ' POR DEFAULT
    FXWSKSEND09.txtUserName = UNAME$
  PWORD$ = SUBTEX$(TTXX$, "PASSWORD=", Chr$(124))                   ' DEL ARCHIVO TRANSFERENCIA
    If PWORD$ = "" Then PWORD$ = Mid$(TTYY$, 513, 32)               ' DE LA CONFIGURACION BASICA
    PWORD$ = TRIM$(DESENCRI$(PWORD$))
    If PWORD$ = "" Then PWORD$ = "123456"                           ' POR DEFAULT
    FXWSKSEND09.txtPassword = PWORD$
    '
  If MAI_AUTO% < 1 Then
    If Len(TTYY$) >= 552 Then
      If TRIM$(Mid$(TTYY$, 545, 8)) <> "PREVIEW" Then
        MAI_AUTO% = 1
      End If
    End If
  End If
  '
  If MAI_AUTO > 0 Then
        Call FXWSKSEND09.cmdSend_Click    ' ENVIO AUTOMATICO
     Else
        FXWSKSEND09.Show 1
  End If
  '
End Sub

