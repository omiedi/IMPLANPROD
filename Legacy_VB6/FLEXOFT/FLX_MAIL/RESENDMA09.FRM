VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form RESENDMA09 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Re-Envío de Mail"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Save"
      Height          =   255
      Left            =   5280
      TabIndex        =   46
      Top             =   2160
      Width           =   495
   End
   Begin VB.ListBox List2 
      Height          =   645
      Left            =   6000
      TabIndex        =   42
      Top             =   2640
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   1560
      TabIndex        =   41
      Top             =   6240
      Width           =   6015
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Control Rechazos por Cuenta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   40
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   840
      Top             =   4800
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Auto"
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
      Left            =   6960
      TabIndex        =   39
      Top             =   5570
      Width           =   615
   End
   Begin VB.Frame fraOptions 
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2100
      Left            =   6000
      TabIndex        =   29
      Top             =   3400
      Width           =   1575
      Begin VB.CheckBox ckPopLogin 
         Caption         =   "POP Login"
         Height          =   250
         Left            =   180
         TabIndex        =   36
         ToolTipText     =   "Use Login Authorization When Connecting to a Host"
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CheckBox ckReceipt 
         Caption         =   "Receipt"
         Height          =   250
         Left            =   180
         TabIndex        =   35
         ToolTipText     =   "Request a Return Receipt"
         Top             =   1365
         Width           =   1035
      End
      Begin VB.ComboBox cboPriority 
         Height          =   315
         ItemData        =   "RESENDMA09.frx":0000
         Left            =   120
         List            =   "RESENDMA09.frx":0002
         TabIndex        =   34
         Text            =   "cboPriority"
         ToolTipText     =   "Sets the Prioirty of the Mail Message"
         Top             =   720
         Width           =   1340
      End
      Begin VB.CheckBox ckHtml 
         Caption         =   "Html"
         Height          =   250
         Left            =   180
         TabIndex        =   33
         ToolTipText     =   "Mail Body is HTML / Plain Text"
         Top             =   1140
         Width           =   1035
      End
      Begin VB.CheckBox ckLogin 
         Caption         =   "Login"
         Height          =   250
         Left            =   180
         TabIndex        =   32
         ToolTipText     =   "Use Login Authorization When Connecting to a Host"
         Top             =   1575
         Width           =   915
      End
      Begin VB.OptionButton optEncodeType 
         Caption         =   "MIME"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   31
         ToolTipText     =   "Use MIME encoding for Mail & Attachments."
         Top             =   240
         Value           =   -1  'True
         Width           =   915
      End
      Begin VB.OptionButton optEncodeType 
         Caption         =   "UUEncode"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   30
         ToolTipText     =   "Use UU Encoding for Attachments."
         Top             =   450
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Change"
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
      Left            =   6800
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Cambiar Cuenta de Envío"
      Top             =   240
      Width           =   780
   End
   Begin VB.TextBox txtServer 
      Height          =   285
      Left            =   1575
      TabIndex        =   24
      Top             =   180
      Width           =   4200
   End
   Begin VB.TextBox txtUserName 
      Height          =   285
      Left            =   1575
      TabIndex        =   23
      Top             =   550
      Width           =   4200
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   6000
      PasswordChar    =   "*"
      TabIndex        =   22
      Top             =   550
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Re-Send"
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
      Left            =   6000
      TabIndex        =   21
      Top             =   5570
      Width           =   950
   End
   Begin VB.TextBox txtToName 
      Height          =   285
      Left            =   1575
      TabIndex        =   10
      Top             =   1800
      Width           =   4200
   End
   Begin VB.TextBox txtTo 
      Height          =   285
      Left            =   1575
      TabIndex        =   9
      Top             =   2160
      Width           =   4200
   End
   Begin VB.TextBox txtFromName 
      Height          =   285
      Left            =   1575
      TabIndex        =   8
      Top             =   1080
      Width           =   4200
   End
   Begin VB.TextBox txtFrom 
      Height          =   285
      Left            =   1575
      TabIndex        =   7
      Top             =   1440
      Width           =   4200
   End
   Begin VB.TextBox txtSubject 
      Height          =   285
      Left            =   1575
      TabIndex        =   6
      Top             =   3600
      Width           =   4200
   End
   Begin VB.TextBox txtMsg 
      Height          =   1620
      Left            =   1575
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   3960
      Width           =   4200
   End
   Begin VB.TextBox txtAttach 
      Height          =   285
      Left            =   1575
      TabIndex        =   4
      Top             =   5640
      Width           =   4200
   End
   Begin VB.TextBox txtCc 
      Height          =   285
      Left            =   1575
      TabIndex        =   3
      Top             =   2880
      Width           =   4200
   End
   Begin VB.TextBox txtCcName 
      Height          =   285
      Left            =   1575
      TabIndex        =   2
      Top             =   2520
      Width           =   4200
   End
   Begin VB.TextBox txtBcc 
      Height          =   285
      Left            =   1575
      TabIndex        =   1
      Top             =   3240
      Width           =   4200
   End
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   6000
      Pattern         =   "*.tmp"
      TabIndex        =   0
      Top             =   960
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog cmDialog 
      Left            =   225
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label DIRELVI 
      Height          =   255
      Left            =   1440
      TabIndex        =   47
      Top             =   2400
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   7560
      TabIndex        =   45
      Top             =   2115
      Width           =   255
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Next:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6165
      TabIndex        =   44
      Top             =   2400
      Width           =   465
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Label5"
      Height          =   255
      Left            =   6600
      TabIndex        =   43
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6720
      TabIndex        =   38
      Top             =   2115
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6120
      TabIndex        =   37
      Top             =   2160
      Width           =   510
   End
   Begin VB.Label lblServer 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SMTP Server:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   0
      TabIndex        =   27
      Top             =   240
      Width           =   1500
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   0
      TabIndex        =   26
      Top             =   660
      Width           =   1500
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5760
      TabIndex        =   25
      Top             =   360
      Width           =   975
   End
   Begin VB.Label lblToName 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Recipient Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   -120
      TabIndex        =   20
      Top             =   1860
      Width           =   1500
   End
   Begin VB.Label lblTo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Recipient Email"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   -120
      TabIndex        =   19
      Top             =   2220
      Width           =   1500
   End
   Begin VB.Label lblFromName 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sender Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   -120
      TabIndex        =   18
      Top             =   1140
      Width           =   1500
   End
   Begin VB.Label lblFrom 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sender Email"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   -120
      TabIndex        =   17
      Top             =   1500
      Width           =   1500
   End
   Begin VB.Label lblSubject 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subject"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   -120
      TabIndex        =   16
      Top             =   3600
      Width           =   1500
   End
   Begin VB.Label lblMsg 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Message"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   -120
      TabIndex        =   15
      Top             =   3960
      Width           =   1500
   End
   Begin VB.Label lblAttach 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Attachment"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   -120
      TabIndex        =   14
      Top             =   5700
      Width           =   1500
   End
   Begin VB.Label lblCC 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cc: Email"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   -120
      TabIndex        =   13
      Top             =   2880
      Width           =   1500
   End
   Begin VB.Label lblCcName 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cc: Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   -120
      TabIndex        =   12
      Top             =   2520
      Width           =   1500
   End
   Begin VB.Label lblBcc 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bcc: Email"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   -120
      TabIndex        =   11
      Top             =   3300
      Width           =   1500
   End
End
Attribute VB_Name = "RESENDMA09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LUCOMAI$

Private Sub Command1_Click()
   '
   Command1.Enabled = False
   '
   MAI_SSMTP$ = txtServer                       ' SMTP
   MAI_SNAME$ = txtFromName                     ' SENDER NAME
   MAI_SMAIL$ = txtFrom                         ' SENDER EMAIL
   MAI_UNAME$ = txtUserName                     ' CUENTA DE USUARIO
   MAI_PWORD$ = txtPassword                     ' CLAVE
   '
   MAI_DESTI$ = txtToName
   MAI_DIREL$ = txtTo
   MAI_ASUNT$ = txtSubject
   MAI_TEXTO$ = txtMsg
   MAI_FTEXT = "PLAIN"
      If ckHtml.Value = 1 Then MAI_FTEXT$ = "HTML"
   MAI_ADJUN$(0) = txtAttach                    ' ADJUNTOS
   '
   CLASEN$ = UCase$(File1.filename)
   MODENVI$ = ""
      If Timer2.Enabled = True Then MODENVI$ = "/AUTO"
   Call REPLATEXT(CLASEN$, ".TMP", "")
   Label5 = CLASEN$
   Call GENERAMAIL(MODENVI$, CLASEN$)
   File1.Refresh
   Label4 = File1.ListCount
   '
   Command1.Enabled = True
   '
 End Sub

Private Sub Command2_Click()
   Timer2.Enabled = True
End Sub

Private Sub Command3_Click()
   '
   Dim CUEMA$(128), VECES%(128)
   For U& = 1 To File1.ListCount
      XX$ = LOADFILE$(LUCOMAI$ + File1.List(U& - 1))
      SMX$ = BUSCAVALOR(XX$, "SENDERMAIL=")
      For KKI% = 1 To CANCUE%
         If CUEMA$(KKI%) = SMX$ Then
            VECES%(KKI%) = VECES%(KKI%) + 1
            GoTo 19
         End If
      Next KKI%
      CANCUE% = CANCUE% + 1
      CUEMA$(CANCUE%) = SMX$
      VECES%(CANCUE%) = 1
19 Next U&
   '
   List1.Clear
   For KKI% = 1 To CANCUE%
      List1.AddItem VECES%(KKI%) & "   " & CUEMA$(KKI%)
   Next KKI%
   Height = 8500
   OPENFORMS = DoEvents
   '
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   DIRELVIE$ = TRIM$(DIRELVI.Caption)
   direlnu$ = TRIM$(txtTo.TEXT)
   If DIRELVIE$ <> "" Then
     If direlnu$ <> DIRELVIE$ Then
       Screen.MousePointer = 11
       For U& = 1 To ULTREG&("LIDIRMAI")
         RELE$ = REGLEIDO$("LIDIRMAI", U&)
         If TRIM$(Mid$(RELE$, 105, 60)) = DIRELVIE$ Then
            Call REPLA(RELE$, direlnu$, 105, 60)
            Call GRAREG("LIDIRMAI", RELE$, U&)
            ATRAMA$ = LUCOMAI$ + File1.filename
            PPXX% = InStr(ATRAMA$, "."): If PPXX% < 1 Then PPXX% = 1 + Len(ATRAMA$)
            On Error Resume Next
            Kill Left$(ATRAMA$, PPXX% - 1) + ".*"
            On Error GoTo 0
            File1.Refresh
            GoTo 99
         End If
       Next U&
     End If
   End If
   '
99 Screen.MousePointer = 1
   Command4.Enabled = True
   '
End Sub

Private Sub Command7_Click()
    '
    CUENTACT& = 0
    For KKI& = 1 To ULTREG&("LISMAIL")
       XX$ = REGLEIDO$("LISMAIL", KKI&)
       If TRIM$(Mid$(XX$, 1, 64)) = txtFrom.TEXT Then
          CUENTACT& = KKI&
          Exit For
       End If
    Next KKI&
    '
    CUENTA& = CUENTACT& + 1
    CANVUEL& = 0
10  If CUENTA& > ULTREG&("LISMAIL") Then CUENTA& = 1
    '
    XX$ = REGLEIDO$("LISMAIL", CUENTA&)
    txtFrom.TEXT = TRIM$(Mid$(XX$, 1, 64)) 'MAIL
    txtUserName.TEXT = TRIM$(Mid$(XX$, 65, 32))
    txtPassword.TEXT = TRIM$(Mid$(XX$, 97, 16))
    txtServer.TEXT = TRIM$(Mid$(XX$, 113, 32))
    STATCUE% = CVI(Mid$(XX$, 145, 2))
    If STATCUE% < 1 Then
       CUENTA& = CUENTA& + 1
       CANVUEL& = CANVUEL& + 1
       If CANVUEL& < 64 Then GoTo 10
    End If
    '
End Sub

Private Sub Form_Load()
    '
    LUCOMAI$ = LUCOM$
    '
    On Error Resume Next
    MkDir LUCOM$ + "MAILCTRL"
    On Error GoTo 0
    If EXISTE%(LUCOM$ + "MAILCTRL\.") > 0 Then
       LUCOMAI$ = LUCOM$ + "MAILCTRL\"
    End If
    If EXISTE%("F:\FLEXOFT\MAILCTRL\.") > 0 Then
       LUCOMAI$ = "F:\FLEXOFT\MAILCTRL\"
    End If
    '
    File1.Path = LUCOMAI$
    File1.Refresh
    Label4 = File1.ListCount
    '
    txtServer.TEXT = Control("", "SMTPSRVR") 'smtpserver
    txtFrom.TEXT = Control("", "FROMMAIL")
    txtUserName.TEXT = Control("", "USUARIO")
    txtPassword.TEXT = Control("", "KLAVE")
    '
    Call Command7_Click      ' para pasar a la siguiente habilitada
    '
End Sub

Sub file1_click()
    '
    ATRAMAI$ = LUCOMAI$ + File1.filename
    If EXISTE%(ATRAMAI$) > 0 Then
         TTXX$ = LOADFILE$(ATRAMAI$)
         Call CARGAENVIO(TTXX$)
       Else
         MsgBox "No Existe " + ATRAMAI$
    End If
    '
End Sub

Sub CARGAENVIO(TTXX$)

  DESTI$ = SUBTEX$(TTXX$, "DESTINO=", Chr$(124))
    RESENDMA09.txtToName.TEXT = DESTI$
  DIREL$ = SUBTEX$(TTXX$, "DIRELEC=", Chr$(124))
    If MAI_TEST > 0 Then DIREL$ = "asistec@flexoft.com.ar"
    RESENDMA09.txtTo.TEXT = DIREL$
    RESENDMA09.DIRELVI.Caption = DIREL$
  ASUNT$ = SUBTEX$(TTXX$, "ASUNTOM=", Chr$(124))
    RESENDMA09.txtSubject.TEXT = ASUNT$
  TEXTOX$ = SUBTEX$(TTXX$, "TEXTOMA=", Chr$(124))
    If Left$(TEXTOX$, 3) = "###" Then                ' ES LA REFERENCIA A UN ARCHIVO CON RUTA
       ATEX$ = TRIM$(Mid$(TEXTOX$, 4))               ' COMPLETA
       If EXISTE%(ATEX$) > 0 Then
          TEXTOX$ = LOADFILE$(ATEX$)
       End If
    End If
    RESENDMA09.txtMsg = TEXTOX$
  FORTX$ = SUBTEX$(TTXX$, "FORTEXT=", Chr$(124))
    If FORTX$ = "HTML" Then RESENDMA09.ckHtml.Value = 1
  ADJUN$ = SUBTEX$(TTXX$, "ADJUNTO=", Chr$(124))
10    PPXX% = InStr(ADJUN$, "&")
      If PPXX% > 0 Then ADJUN$ = Left$(ADJUN$, PPXX% - 1) + ";" + Mid$(ADJUN$, PPXX% + 1): GoTo 10
      PPXX% = InStr(ADJUN$, " ;")
      If PPXX% > 0 Then ADJUN$ = Left$(ADJUN$, PPXX% - 1) + ";" + Mid$(ADJUN$, PPXX% + 2): GoTo 10
      PPXX% = InStr(ADJUN$, "; ")
      If PPXX% > 0 Then ADJUN$ = Left$(ADJUN$, PPXX% - 1) + ";" + Mid$(ADJUN$, PPXX% + 2): GoTo 10
    RESENDMA09.txtAttach.TEXT = ADJUN$
  SOLCO$ = SUBTEX$(TTXX$, "SOLCONF=", Chr$(124))
    If SOLCO$ = "SI" Then RESENDMA09.ckReceipt.Value = 1
  RESENDMA09.ckLogin.Value = 1
  '
  TTYY$ = LOADFILE(LUCOM$ + "STPFMAIL.CFG")
  SNAME$ = SUBTEX$(TTXX$, "SENDERNAME=", Chr$(124))                 ' DEL ARCHIVO TRANSFERENCIA
    If SNAME$ = "" Then SNAME$ = TRIM$(Mid$(TTYY$, 129, 128))       ' DE LA CONFIGURACION BASICA
    If SNAME$ = "" Then SNAME$ = "FLEXOFT Soluciones"               ' POR DEFAULT
    RESENDMA09.txtFromName = SNAME$
    '
End Sub

Private Sub Timer2_Timer()
   '
   Static SEGDESULTEN&, ULTIRE&, ENVIANDO%, ULTIAR$
   '
   If ENVIANDO% <> 1 Then
      File1.Refresh
      Label4 = File1.ListCount
   End If
   '
   SEGDESULTEN& = SEGDESULTEN& + 1
   Label7 = SEGDESULTEN&: Label7.Refresh
10 If SEGDESULTEN& < 10 Then
     ATRAMAI$ = LUCOMAI$ + ULTIAR$
     If EXISTE%(ATRAMAI$) > 0 Then Exit Sub        ' DEPUES AJUSTAR AL INTERVALO
   End If
   SEGDESULTEN& = 0
   Label7 = SEGDESULTEN&: Label7.Refresh
   '
   File1.Refresh
   Label4 = File1.ListCount
   If File1.ListCount < 10 Then
      Unload Me
      Exit Sub
   End If
   '
   ENVIANDO% = 1
20 File1.ListIndex = ULTIRE&
   File1.Selected(ULTIRE&) = True
   FINAM$ = File1.filename
   If FINAM$ = ULTIAR$ Then
      ULTIRE& = ULTIRE& + 1
      If ULTIRE& > File1.ListCount - 1 Then ULTIRE& = 0
      GoTo 20
   End If
   List2.AddItem FINAM$
   ULTIAR$ = FINAM$
   ATRAMAI$ = LUCOMAI$ + FINAM$
   If EXISTE%(ATRAMAI$) > 0 Then
       TTXX$ = LOADFILE$(ATRAMAI$)
       Call CARGAENVIO(TTXX$)
       Call Command1_Click
       Call Command7_Click     ' para cambiar cuenta de envio
   End If
   ENVIANDO% = 0
   '
End Sub
