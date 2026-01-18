VERSION 5.00
Begin VB.Form TRALECAR 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Traducción Leyendas y Carteles"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Explorando Leyendas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      Left            =   1365
      TabIndex        =   19
      Top             =   2940
      Visible         =   0   'False
      Width           =   2745
      Begin VB.FileListBox File1 
         Height          =   870
         Left            =   315
         TabIndex        =   20
         Top             =   945
         Width           =   2115
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   315
         TabIndex        =   21
         Top             =   420
         Width           =   2115
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "View"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   11160
      Picture         =   "TRALECAR.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Explora Leyendas en Formularios"
      Top             =   1365
      Width           =   645
   End
   Begin VB.Frame Frame2 
      Caption         =   "Portugués"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Index           =   5
      Left            =   5565
      TabIndex        =   17
      Top             =   5880
      Width           =   5475
      Begin VB.CommandButton Command1 
         Height          =   540
         Index           =   5
         Left            =   4830
         Picture         =   "TRALECAR.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   420
         Width           =   540
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   5
         Left            =   105
         MultiLine       =   -1  'True
         TabIndex        =   24
         Top             =   420
         Width           =   4635
      End
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
      Height          =   960
      Left            =   11160
      Picture         =   "TRALECAR.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   210
      Width           =   645
   End
   Begin VB.Frame Frame2 
      Caption         =   "Italiano"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Index           =   4
      Left            =   5565
      TabIndex        =   9
      Top             =   4725
      Width           =   5475
      Begin VB.CommandButton Command1 
         Height          =   540
         Index           =   4
         Left            =   4830
         Picture         =   "TRALECAR.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   420
         Width           =   540
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   4
         Left            =   105
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   420
         Width           =   4635
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Francais"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Index           =   3
      Left            =   5565
      TabIndex        =   7
      Top             =   3570
      Width           =   5475
      Begin VB.CommandButton Command1 
         Height          =   540
         Index           =   3
         Left            =   4830
         Picture         =   "TRALECAR.frx":0A68
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   420
         Width           =   540
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   3
         Left            =   105
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   420
         Width           =   4635
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Deutsch"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Index           =   2
      Left            =   5565
      TabIndex        =   5
      Top             =   2415
      Width           =   5475
      Begin VB.CommandButton Command1 
         Height          =   540
         Index           =   2
         Left            =   4830
         Picture         =   "TRALECAR.frx":0D72
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   420
         Width           =   540
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   2
         Left            =   105
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   420
         Width           =   4635
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "English"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Index           =   1
      Left            =   5565
      TabIndex        =   3
      Top             =   1260
      Width           =   5475
      Begin VB.CommandButton Command1 
         Height          =   540
         Index           =   1
         Left            =   4830
         Picture         =   "TRALECAR.frx":107C
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   420
         Width           =   540
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   1
         Left            =   105
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   420
         Width           =   4635
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Español"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Index           =   0
      Left            =   5565
      TabIndex        =   2
      Top             =   105
      Width           =   5475
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
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
         Height          =   540
         Left            =   105
         TabIndex        =   16
         Top             =   420
         Width           =   5265
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Indice"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6840
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   5265
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         TabIndex        =   23
         Text            =   "Text2"
         Top             =   6350
         Width           =   4320
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5715
         Left            =   105
         TabIndex        =   1
         Top             =   420
         Width           =   5055
      End
      Begin VB.Label Label3 
         Caption         =   "Buscar:"
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
         TabIndex        =   22
         Top             =   6405
         Width           =   960
      End
   End
End
Attribute VB_Name = "TRALECAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FINUM%(16)
Dim BUFTRA As String * 128

Private Sub Command1_Click(Index As Integer)
   '
   Static FITRA%
   Static BUFTRA As String * 1024
   Static URETRA&
   '
   If FITRA% < 1 Then
     Screen.MousePointer = 11
     FITRA% = FreeFile
     Open LUBAS$ + "TRALABS.DRV" For Random Access Read Write Shared As #FITRA% Len = 1024
     URETRA& = LOF(FITRA%) / 1024
     Screen.MousePointer = 1
   End If
   '
   If List1.ListIndex >= 0 Then
     REIN& = 1 + List1.ListIndex
     KKI% = Index
     '
     Get #FITRA%, REIN&, BUFTRA$
     TTXX$ = BUFTRA$
     Call REPLA(TTXX$, Text1(KKI%), 128 * KKI% + 1, 128)
     LSet BUFTRA$ = TTXX$
     Put #FITRA%, REIN&, BUFTRA$
   End If
   '
End Sub

Private Sub Command2_Click()
   Close
   End
End Sub

Private Sub Command3_Click()
   '
   Static FITRA%
   Static BUFTRA As String * 1024
   Static URETRA&
   Static XINDI$
   '
   Frame3.Visible = True
   Screen.MousePointer = 11
   File1.Path = "C:\Archivos de programa\DevStudio\Vb\Flexoft\Formularios"
5  File1.Pattern = "*.FRM"
   '
   If FITRA% < 1 Then
     Screen.MousePointer = 11
     FITRA% = FreeFile
     Open LUBAS$ + "TRALABS.DRV" For Random Access Read Write Shared As #FITRA% Len = 1024
     URETRA& = LOF(FITRA%) / 1024
     XINDI$ = String$(2 * URETRA&, 0)
     For U& = 1 To URETRA&
       Get #FITRA%, U&, BUFTRA$
       Mid$(XINDI$, 2 * U& - 1, 2) = Left$(BUFTRA$, 2)
     Next U&
     Screen.MousePointer = 1
   End If
   '
   For KL& = 1 To File1.ListCount
     Label2 = File1.List(KL& - 1)
     Label2.Refresh
     AFUNA$ = File1.Path + "\" + File1.List(KL& - 1)
     '
     If FileLen(AFUNA$) > 0 Then
       NX% = FreeFile
       Open AFUNA$ For Input As #NX%
       '
       While Not EOF(NX%)
         Line Input #NX%, AA$
         BB$ = TRIM$(AA$)
         '
         PPXX% = InStr(UCase(BB$), "CAPTION")
         If PPXX% > 0 Then
           CC$ = TRIM$(Mid$(BB$, PPXX% + 12))
           If Left$(CC$, 1) = "=" Then
             DD$ = TRIM$(Mid$(CC$, 2))
             If Left$(DD$, 1) = Chr$(34) Then
               DK$ = TRIM$(Mid$(DD$, 2))
               LONI% = Len(DK$)
               If Mid$(DK$, LONI%, 1) = Chr$(34) Then
                 TESPA$ = TRIM$(Left$(DK$, LONI% - 1))
                 '
                 GoSub 50
                 GoTo 19
                 '
               End If
             End If
           End If
         End If
         '
         PPXX% = InStr(UCase(BB$), "TOOLTIPTEXT")
         If PPXX% > 0 Then
           CC$ = TRIM$(Mid$(BB$, PPXX% + 11))
           If Left$(CC$, 1) = "=" Then
             DD$ = TRIM$(Mid$(CC$, 2))
             If Left$(DD$, 1) = Chr$(34) Then
               DK$ = TRIM$(Mid$(DD$, 2))
               LONI% = Len(DK$)
               If Mid$(DK$, LONI%, 1) = Chr$(34) Then
                 TESPA$ = TRIM$(Left$(DK$, LONI% - 1))
                 '
                 GoSub 50
                 GoTo 19
                 '
               End If
             End If
           End If
         End If
         '
19     Wend
       '
       Close #NX%
     End If
     '
   Next KL&
   '
   If File1.Path = "C:\Archivos de programa\DevStudio\Vb\Flexoft\Formularios" Then
      File1.Path = "C:\Archivos de programa\DevStudio\Vb\Flexoft\Form_ACC"
      GoTo 5
   End If
   '
   Close #FITRA%: FITRA% = 0
   Call SAVEFILE(LUBAS$ + "TRALABS.X01", XINDI$)
   '
   Frame3.Visible = False
   Screen.MousePointer = 1
   '
   Screen.MousePointer = 11
   FITRA% = FreeFile
   Open LUBAS$ + "TRALABS.DRV" For Random Access Read Write Shared As #FITRA% Len = 1024
   List1.Clear
   For U& = 1 To URETRA&
     Get #FITRA%, U&, BUFTRA$
     List1.AddItem Mid$(BUFTRA$, 3, 126)
   Next U&
   Close #FITRA%: FITRA% = 0
   Screen.MousePointer = 1
   '
   Exit Sub
   '
50 BVT% = BINVAL%(TESPA$)
   AZ$ = MKI$(BVT%)
   RFF$ = "": PINI& = 1
60 POSI& = InStr(PINI&, XINDI$, AZ$)
   If POSI& < 1 Then GoTo 70                           ' NEXT RGX&
   If (POSI& Mod 2) <> 1 Then
      PINI& = POSI& + 1
      GoTo 60
   End If
   PINI& = POSI& + 2
   RGL& = (POSI& + 1) / 2
   RFF$ = RFF$ + MKS$(RGL&)
   GoTo 60
   '
70 For KKI% = 1 To Len(RFF$) Step 4
     REIN& = CVS(Mid$(RFF$, KKI%, 4))
     If REIN& > 0 Then
       If REIN& <= URETRA& Then
         Get #FITRA%, REIN&, BUFTRA$
         TINDI$ = TRIM$(Mid$(BUFTRA$, 3, 126))
         If TINDI$ = TESPA$ Then GoTo 90
       End If
     End If
   Next KKI%
   '
   TTXX$ = MKI$(0) + Space$(1022)
   Call REPLA(TTXX$, MKI$(BVT%), 1, 2)
   Call REPLA(TTXX$, TESPA$ + Space$(128), 3, 126)
   BUFTRA$ = TTXX$
   URETRA& = URETRA& + 1
   Put #FITRA%, URETRA&, BUFTRA$
   XINDI$ = XINDI$ + MKI$(BVT%)
   REIN& = URETRA&
   '
90 Return
   '
End Sub


Private Sub Form_Load()
   'Show
   'Call VERJOBID(OKEY%)
   '
   Screen.MousePointer = 11
   FITRA% = FreeFile
   Open LUBAS$ + "TRALABS.DRV" For Random Access Read Write Shared As #FITRA% Len = 1024
   URETRA& = LOF(FITRA%) / 1024
   List1.Clear
   For U& = 1 To URETRA&
     Get #FITRA%, U&, BUFTRA$
     List1.AddItem Mid$(BUFTRA$, 3, 126)
   Next U&
   Close #FITRA%: FITRA% = 0
   Screen.MousePointer = 1
   '
End Sub

Private Sub List1_Click()
   '
   Static FITRA%
   Static BUFTRA As String * 1024
   Static URETRA&
   '
   If FITRA% < 1 Then
     Screen.MousePointer = 11
     FITRA% = FreeFile
     Open LUBAS$ + "TRALABS.DRV" For Random Access Read Write Shared As #FITRA% Len = 1024
     URETRA& = LOF(FITRA%) / 1024
     Screen.MousePointer = 1
   End If
   '
   Label1 = List1.TEXT
   REIN& = 1 + List1.ListIndex
   '
   Get #FITRA%, REIN&, BUFTRA$
   For KKI% = 1 To 5
     Text1(KKI%) = TRIM$(Mid$(BUFTRA$, 128 * KKI% + 1, 128))
     Text1(KKI%).Refresh
   Next KKI%
   '
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     vuelta% = 0
     ABUS$ = UCase$(TRIM$(Text2))
10   For KKI% = List1.ListIndex + 1 To List1.ListCount
       If InStr(UCase$(List1.List(KKI% - 1)), ABUS$) > 0 Then
         List1.ListIndex = KKI% - 1
         If List1.TopIndex < List1.ListIndex - 25 Or List1.TopIndex > List1.ListIndex Then
           On Error Resume Next
           List1.TopIndex = List1.ListIndex - 20
           On Error GoTo 0
           Exit Sub
         End If
       End If
     Next KKI%
     If vuelta% = 0 Then
       vuelta% = vuelta% + 1
       List1.ListIndex = 0
       GoTo 10
     End If
   End If
End Sub
