VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form FRPEDREM 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recepcion Pedidos Remotos"
   ClientHeight    =   1050
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   3735
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1050
   ScaleWidth      =   3735
   StartUpPosition =   2  'CenterScreen
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   2625
      Top             =   -210
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
      Left            =   1680
      Top             =   -210
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   327680
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Setup"
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
      Left            =   1200
      Picture         =   "FRPEDREM.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Configuración de Funcionamiento"
      Top             =   210
      Width           =   750
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   105
      ScaleHeight     =   75
      ScaleWidth      =   900
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   210
      Width           =   960
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.CommandButton Command1 
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
      Height          =   690
      Left            =   2880
      Picture         =   "FRPEDREM.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Terminar - Salir de la Aplicación"
      Top             =   210
      Width           =   750
   End
   Begin VB.CommandButton Command29 
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
      Height          =   690
      Left            =   2040
      Picture         =   "FRPEDREM.frx":058C
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Ayuda Flexoft en Línea"
      Top             =   210
      Width           =   750
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   -380
      Picture         =   "FRPEDREM.frx":0896
      Top             =   390
      Width           =   2010
   End
End
Attribute VB_Name = "FRPEDREM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Call FINAL("")
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   Call CONECRUN("FLSETUP", "Configuración de Funcionamiento")
   Command2.Enabled = True
End Sub

Private Sub RECIBEPEDS()
   '
   VERDAFAL = False
   If COMALTER("Desea Refrescar B.Datos Mails Recibidos ?\\Si,Refrescar\No, Continuar") = 1 Then
     VERDAFAL = True
   End If
   '
   Screen.MousePointer = 11
   MAPISession1.DownLoadMail = VERDAFAL
   MAPISession1.LogonUI = False
   MAPISession1.Action = 1
   MAPIMessages1.SessionID = MAPISession1.SessionID
   '
   DoEvents
   MAPIMessages1.Fetch
   DoEvents
   CAMENSA& = MAPIMessages1.MsgCount
   DoEvents
   '
   For U& = 1 To CAMENSA&
     Call TRACE(20, U&, CAMENSA&)
10   If U& > MAPIMessages1.MsgCount Then GoTo 99
     MAPIMessages1.MsgIndex = U& - 1
     ASUNTO$ = TRIM$(MAPIMessages1.MsgSubject)
     If Left$(ASUNTO$, 7) = "P-Elec." Then
       CRLF$ = Chr$(13) + Chr$(10)
       TEMAI$ = MAPIMessages1.MsgNoteText + CRLF$
       While TEMAI$ <> ""
         PPXX% = InStr(TEMAI$, CRLF$)
         If PPXX% < 1 Then GoTo 90
         RENGLX$ = TRIM$(Left$(TEMAI$, PPXX% - 1))
         TEMAI$ = Mid$(TEMAI$, PPXX% + 2)
         If RENGLX$ <> "" Then
           IDREN$ = Left$(RENGLX$, 8)
           If IDREN$ = "N-CLIEN=" Then
               NCLIE$ = TRIM$(Mid$(RENGLX$, 9))
             ElseIf IDREN$ = "D-CLIEN=" Then
               DCLIE$ = TRIM$(Mid$(RENGLX$, 9))
             ElseIf IDREN$ = "F-PEDID=" Then
               FPEDI$ = TRIM$(Mid$(RENGLX$, 9))
             ElseIf IDREN$ = "N-SUCUR=" Then
               NSUCU$ = TRIM$(Mid$(RENGLX$, 9))
             ElseIf IDREN$ = "L-ENTRE=" Then
               LENTR$ = TRIM$(Mid$(RENGLX$, 9))
             ElseIf IDREN$ = "T-OBSER=" Then
               OBSGE$ = TRIM$(Mid$(RENGLX$, 9))
             Else
               CODD$ = TRIM$(Left$(RENGLX$, 16))
               CANX$ = TRIM$(Mid$(RENGLX$, 17, 8))
               FENT$ = TRIM$(Mid$(RENGLX$, 25, 8))
' AQUI AGREGAR A ARCHIVO
           End If
         End If
       Wend
90         Screen.MousePointer = 11
           If ATENTI% = 1 Then
             MAPIMessages1.Delete
             GoTo 10
           End If
     End If
   Next U&
   '
99 On Error Resume Next
   MAPISession1.Action = 2
   On Error GoTo 0
   '
   Screen.MousePointer = 1
End Sub

Private Sub Form_Load()
   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then Call FINAL("")
   '
   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
     Caption = Caption + "  -  " + EPAC$
   End If
   '
   Call RECIBEPEDS
End Sub

Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

