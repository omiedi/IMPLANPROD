VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form FORMAIL 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Circulares por Mail"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   8460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   735
      Top             =   0
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
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   327680
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   7560
      Picture         =   "FORMAIL.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Enviar"
      Top             =   945
      Width           =   750
   End
   Begin VB.Frame Frame2 
      Caption         =   "Presentación Mensaje Personalizado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4845
      Left            =   105
      TabIndex        =   3
      Top             =   105
      Width           =   7260
      Begin VB.TextBox Text3 
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
         Left            =   1050
         TabIndex        =   6
         Text            =   "Text3"
         Top             =   4365
         Width           =   6105
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3900
         Left            =   105
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Text            =   "FORMAIL.frx":030A
         Top             =   315
         Width           =   7050
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Dir.Elec:"
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
         TabIndex        =   5
         Top             =   4410
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mensaje Circular"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3165
      Left            =   105
      TabIndex        =   1
      Top             =   5040
      Width           =   7260
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2640
         Left            =   105
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Text            =   "FORMAIL.frx":0310
         Top             =   315
         Width           =   7050
      End
   End
   Begin VB.CommandButton Command2 
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   7560
      Picture         =   "FORMAIL.frx":05B0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   210
      Width           =   750
   End
End
Attribute VB_Name = "FORMAIL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Dim DESTIX$(2)
   '
   MAPISession1.DownLoadMail = False
   MAPISession1.LogonUI = False
   MAPISession1.Action = 1
   MAPIMessages1.SessionID = MAPISession1.SessionID
   '
   For U& = 1 To 5 'ULTREG&("BSEGCOM")
   XX$ = REGLEIDO$("BSEGCOM", U&)
     DIREC$ = TRIM$(Mid$(XX$, 65, 64))
     If DIREC$ <> "" Then
       EMPRE$ = TRIM$(Mid$(XX$, 3, 59))
       CALNUM$ = TRIM$(Mid$(XX$, 129, 48))
       CPOST$ = TRIM$(Mid$(XX$, 177, 8))
       If CPOST$ <> "" Then CPOST$ = CPOST$ + " "
       LOCAX$ = CPOST$ + TRIM$(Mid$(XX$, 185, 32))
       PROPA$ = TRIM$(TRIM$(Mid$(XX$, 217, 32)) + " " + TRIM$(Mid$(XX$, 249, 24)))
       '
       CADEST% = 0
       For KKI% = 0 To 2
         If TRIM$(Mid$(XX$, 449 + 96 * KKI%, 32)) <> "" Then
           CADEST% = CADEST% + 1
           DESTIX$(CADEST%) = TRIM$(Mid$(XX$, 449 + 96 * KKI%, 32))
         End If
       Next KKI%
       '
       TTXX$ = ""
       If EMPRE$ <> "" Then TTXX$ = TTXX$ + EMPRE$ + Chr$(13) + Chr$(10)
       If CALNUM$ <> "" Then TTXX$ = TTXX$ + CALNUM$ + Chr$(13) + Chr$(10)
       If LOCAX$ <> "" Then TTXX$ = TTXX$ + LOCAX$ + Chr$(13) + Chr$(10)
       If PROPA$ <> "" Then TTXX$ = TTXX$ + PROPA$ + Chr$(13) + Chr$(10)
       '
       Text3 = DIREC$
       Text3.Refresh
       '
       If CADEST% < 1 Then
           TTYY$ = "Sres." + Chr$(13) + Chr$(10) + TTXX$
           TTYY$ = TTYY$ + Chr$(13) + Chr$(10)
           TTYY$ = TTYY$ + Chr$(13) + Chr$(10)
           TTYY$ = TTYY$ + Text1
           GoSub SENDMAIL
         Else
           For KKI% = 1 To CADEST%
             TTYY$ = "Sr./Sra." + Chr$(13) + Chr$(10)
             TTYY$ = TTYY$ + DESTIX$(KKI%) + Chr$(13) + Chr$(10)
             TTYY$ = TTYY$ + TTXX$
             TTYY$ = TTYY$ + Chr$(13) + Chr$(10)
             TTYY$ = TTYY$ + Chr$(13) + Chr$(10)
             TTYY$ = TTYY$ + Text1
             GoSub SENDMAIL
           Next KKI%
       End If
       
     End If
   Next U&
   '
   Call CIERRARCH("*.*")
   MAPIMessages1.SessionID = 0
   MAPISession1.Action = 2
   '
   MAPISession1.Action = 1
   MAPIMessages1.SessionID = MAPISession1.SessionID
   MAPIMessages1.Action = 3
   MAPISession1.Action = 2
   '
   Unload Me
   '
   Exit Sub
   '
SENDMAIL:
   '
   Text2 = TTYY$
   Text2.Refresh
   'MAPISession1.Action = 1
   'MAPIMessages1.SessionID = MAPISession1.SessionID
   MAPIMessages1.MsgIndex = -1
   MAPIMessages1.MsgSubject = "cuarta prueba" + Str$(KKI%)
   MAPIMessages1.RecipAddress = "asistec@flexoft.com.ar"
   MAPIMessages1.MsgNoteText = Text2
   'MsgBox Text2
   'MAPIMessages1.Action = 4   ' vbMessageSAVE
   MAPIMessages1.Action = 3   ' vbMessageSend
   Return
   '
End Sub

Private Sub Command2_Click()
   Call CIERRARCH("*.*")
   Hide
End Sub

