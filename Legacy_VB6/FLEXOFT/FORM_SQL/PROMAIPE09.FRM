VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form PROMAIPE09 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Muestra Mensaje de Correo Electrónico Recibido"
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9420
   BeginProperty Font 
      Name            =   "Small Fonts"
      Size            =   6
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   9420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   8640
      Picture         =   "PROMAIPE09.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Borra los Marcados"
      Top             =   1560
      Width           =   645
   End
   Begin VB.Frame Frame3 
      Caption         =   "Pendientes de Procesar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      TabIndex        =   24
      Top             =   120
      Width           =   8415
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2160
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   25
         Top             =   360
         Width           =   8175
      End
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   8650
      ScaleHeight     =   75
      ScaleWidth      =   585
      TabIndex        =   3
      Top             =   2280
      Width           =   650
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.CommandButton Command7 
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
      Height          =   675
      Left            =   8650
      Picture         =   "PROMAIPE09.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Cierra y Abandona la Aplicación"
      Top             =   210
      Width           =   645
   End
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
   Begin VB.Frame Frame2 
      Caption         =   "Detalle del Mensaje Recibido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5370
      Left            =   105
      TabIndex        =   0
      Top             =   2880
      Width           =   9180
      Begin VB.Frame Frame1 
         Caption         =   "Origen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   8955
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Dir-elec:"
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
            Left            =   315
            TabIndex        =   23
            Top             =   720
            Width           =   750
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "De:"
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
            Left            =   210
            TabIndex        =   22
            Top             =   405
            Width           =   855
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            Caption         =   "Asunto:"
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
            TabIndex        =   21
            Top             =   1365
            Width           =   960
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Para:"
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
            Left            =   320
            TabIndex        =   20
            Top             =   1030
            Width           =   750
         End
         Begin VB.Label Label43 
            Alignment       =   1  'Right Justify
            Caption         =   "Cte:"
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
            Left            =   6960
            TabIndex        =   19
            Top             =   360
            Width           =   555
         End
         Begin VB.Label Label26 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   7560
            TabIndex        =   18
            Top             =   300
            Width           =   1245
         End
         Begin VB.Label Label14 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1200
            TabIndex        =   17
            ToolTipText     =   "Doble-Click para Incluir en Regla de Borrado"
            Top             =   300
            Width           =   5685
         End
         Begin VB.Label Label15 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1200
            TabIndex        =   16
            ToolTipText     =   "Doble-Click para Incluir en Regla de Borrado"
            Top             =   1275
            Width           =   6645
         End
         Begin VB.Label Label16 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1200
            TabIndex        =   15
            ToolTipText     =   "Doble-Click para Incluir en Regla de Borrado"
            Top             =   630
            Width           =   5085
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   8040
            TabIndex        =   14
            Top             =   1260
            Width           =   765
         End
         Begin VB.Label Label4 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6270
            TabIndex        =   13
            Top             =   945
            Width           =   2535
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "Id:"
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
            Left            =   5640
            TabIndex        =   12
            Top             =   1005
            Width           =   555
         End
         Begin VB.Label Label17 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1200
            TabIndex        =   11
            Top             =   960
            Width           =   4545
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
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
            Height          =   285
            Left            =   7080
            TabIndex        =   10
            Top             =   630
            Width           =   1725
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Recep:"
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
            Left            =   6360
            TabIndex        =   9
            Top             =   720
            Width           =   675
         End
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   3090
         Left            =   120
         TabIndex        =   7
         Top             =   2160
         Visible         =   0   'False
         Width           =   8895
         ExtentX         =   15690
         ExtentY         =   5450
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3090
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Text            =   "PROMAIPE09.frx":0454
         Top             =   2160
         Visible         =   0   'False
         Width           =   8910
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3090
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Text            =   "PROMAIPE09.frx":045C
         Top             =   2160
         Width           =   8910
      End
      Begin VB.Label HUBORRA 
         Height          =   255
         Left            =   6960
         TabIndex        =   29
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8640
      TabIndex        =   28
      Top             =   960
      Width           =   650
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8640
      TabIndex        =   27
      Top             =   1260
      Width           =   650
   End
   Begin VB.Image Image4 
      Height          =   390
      Left            =   8230
      Picture         =   "PROMAIPE09.frx":0462
      Top             =   2490
      Width           =   1515
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   200
      TabIndex        =   1
      Top             =   1890
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "PROMAIPE09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    '
    Command1.Enabled = False
    Screen.MousePointer = 11
    HUBORRA = "0"
    For KKI% = 1 To List1.ListCount
       If List1.Selected(KKI% - 1) = True Then
          NMAI& = XVALO(Left$(List1.List(KKI% - 1), 6))
          CONDI$ = "NUME_MAIL = " & NMAI&
          Call ACTUREGISTRO("BANPENPRO", "PROS_DEST", CONDI$, "", RAFE1&, "NOMESS")
          Call ACTUREGISTRO("BANPENPRO", "BORR_MAIL", CONDI$, 1, RAFE2&, "NOMESS")
          If RAFE1& > 0 Or RAFE2& > 0 Then HUBORRA = "1"
       End If
    Next KKI%
    Screen.MousePointer = 1
    '
    Command1.Enabled = True
    Hide
    '
End Sub

Private Sub Command7_Click()
   Label1 = "0"
   HUBORRA = "0"
   Hide
End Sub

Private Sub Form_Activate()
   If ATENTI% > 0 Then
     Left = 0
     Top = 0
     ATENTI% = 0
     Text2.SetFocus
   End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UTIMETEC = Timer
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Cancel = 0
End Sub

Private Sub Label13_Change()
    WebBrowser1.Visible = False
    NUMAI& = XVALO(Label13)
    If NUMAI& < 1 Then
          Label26 = ""
          Label17 = ""
          Label4 = ""
          Label15 = ""
          Label16 = ""
          Label14 = ""
          '
          WebBrowser1.Visible = False
          Text1.TEXT = ""
          Exit Sub
    End If
    '
    CONDI$ = "NUME_MAIL = " & NUMAI&
    If CantRegistros("BANPENPRO", CONDI$, "NOMESS") > 0 Then
          '
          Caption = "Muestra Mensaje de Correo Recibido#" & TRIM$(Str$(NUMAI&))
          Label26 = TRIM$(VALOBADA("BANPENPRO", "PROS_DEST", CONDI$, "NOMESS"))
          Label17 = TRIM$(VALOBADA("BANPENPRO", "PARA_MAIL", CONDI$, "NOMESS"))
          Label4 = TRIM$(VALOBADA("BANPENPRO", "IDEN_MAIL", CONDI$, "NOMESS"))
          Label15 = TRIM$(VALOBADA("BANPENPRO", "ASUN_MAIL", CONDI$, "NOMESS"))
          Label16 = TRIM$(VALOBADA("BANPENPRO", "DREL_MAIL", CONDI$, "NOMESS"))
          Label14 = TRIM$(VALOBADA("BANPENPRO", "ORIG_MAIL", CONDI$, "NOMESS"))
          '
          TXMAIL$ = VALOBADA("BANPENPRO", "TEXT_MAIL", CONDI$, "NOMESS")
          HTMAIL$ = VALOBADA("BANPENPRO", "HTMT_MAIL", CONDI$, "NOMESS")
          '
          If TRIM$(HTMAIL$) <> "" Then
               Call SAVEFILE("C:\FLEXOFT\ARADJUN\HTMMAIL.HTM", HTMAIL$)
               WebBrowser1.Navigate "C:\FLEXOFT\ARADJUN\HTMMAIL.HTM"
               WebBrowser1.Visible = True
            Else
               Text1.TEXT = TXMAIL$
               Text1.Refresh
          End If
          '
       Else
          Call MENSERR(24, "Imposible Abrir Mensaje Enviado Numero '" & TRIM$(Str$(NUMAI&)) & "'")
   End If
'    Call ACTUREGISTRO("BANPENPRO", "FHOR_RECEP", CONDI$, FRECEP, RAFE&, "NOMESS")
End Sub

Private Sub Label4_DblClick()
   Label4 = ""
   Text1 = ""
End Sub

Private Sub Label14_DblClick()
    CRITBORR09.Text1(0) = TRIM$(Label14)
    CRITBORR09.Combo1(0) = "Remitente"
    CRITBORR09.Show 1
    Label1 = "2"
    Hide
End Sub

Private Sub Label15_DblClick()
    CRITBORR09.Text1(0) = TRIM$(Label15)
    CRITBORR09.Combo1(0) = "Asunto"
    CRITBORR09.Show 1
    Label1 = "2"
    Hide
End Sub

Private Sub Label16_DblClick()
    CRITBORR09.Text1(0) = TRIM$(Label16)
    CRITBORR09.Combo1(0) = "Remitente"
    CRITBORR09.Show 1
    Label1 = "2"
    Hide
End Sub

Private Sub List1_Click()
    NMAI& = XVALO(Left$(List1.TEXT, 6))
    PROMAIPE09.Label13 = ""
    If NMAI& > 0 Then
       PROMAIPE09.Label13 = TRIM$(Str$(NMAI&))
    End If
    Label10 = List1.SelCount
End Sub

Private Sub List1_DblClick()
    '
    NMAI& = XVALO(Left$(List1.TEXT, 6))
    If NMAI& > 0 Then
       PRINMAI09.Label13 = TRIM$(Str$(NMAI&))
       PRINMAI09.Show 1
       If PRINMAI09.Label1 = "1" Then
          ORIMESS$ = TRIM$(PRINMAI09.Label26)
          If UCase$(Left$(ORIMESS$, 6)) = "DELETE" Then
                CONDI$ = "NUME_MAIL = " & NMAI&
                Call ACTUREGISTRO("BANPENPRO", "PROS_DEST", CONDI$, "", RAFE1&, "NOMESS")
                Call ACTUREGISTRO("BANPENPRO", "BORR_MAIL", CONDI$, 1, RAFE2&, "NOMESS")
                GoTo 90 ' If RAFE1& > 0 Or RAFE2& > 0 Then GoTo 10
              ElseIf ORIMESS$ <> "" Then
                CONDI$ = "NUME_MAIL = " & NMAI&
                DREMIT$ = TRIM$(VALOBADA("BANPENPRO", "DREL_MAIL", CONDI$, "NOMESS"))
                CONDI$ = "DREL_MAIL = '" & DREMIT$ & "'"
                Call ACTUREGISTRO("BANPENPRO", "PROS_DEST", CONDI$, ORIMESS$, RAFE&, "NOMESS")
                Call ACTUREGISTRO("BANPENPRO", "BORR_MAIL", CONDI$, 1, RAFE&, "NOMESS")
                GoTo 90 ' If RAFE& > 0 Then GoTo 10
          End If
          GoTo 90  '          GoTo 11
        ElseIf PRINMAI09.Label1 = 2 Then
          Call BORRAPENDIENTES(OKX%)
          GoTo 90 '          GoTo 10
       End If
    End If
    '
90  CONDIX$ = "NUME_MAIL > 0 "
    CONDIX$ = CONDIX$ + " AND (PROS_DEST = '' OR PROS_DEST IS NULL) "
    CONDIX$ = CONDIX$ + " AND (BORR_MAIL = 0 OR BORR_MAIL IS NULL) "
    CONDIX$ = CONDIX$ + " ORDER BY FHOR_RECEP ASC"
    Call CARGALISTA(PROMAIPE09.List1, "BANPENPRO", "NUME_MAIL/F6; FHOR_RECEP/D17; ASUN_MAIL/A60", CONDIX$)
    Label9 = List1.ListCount
    '
End Sub

Sub BORRAPENDIENTES(OKX%)
    '
    Static YACARLI%
    Static CANCRI%, STRI1$(), UBI1$(), STRI2$(), UBI2$()
    '
    Call CREACAMPO("", "CRITBORR", "NCRITBO", 3, 2, 1)
    Call CREACAMPO("", "CRITBORR", "STRING1", 10, 128)
    Call CREACAMPO("", "CRITBORR", "UBICA1", 10, 32)
    Call CREACAMPO("", "CRITBORR", "STRING2", 10, 128)
    Call CREACAMPO("", "CRITBORR", "UBICA2", 10, 32)
    '
    If CantRegistros&("CRITBORR", "NCRITBO > 0", "NOMESS") <> CANCRI% Then
       YACARLI% = 0
    End If
    '
    If YACARLI% < 1 Then
       Dim CRIBO As ADODB.Recordset
       Set CRIBO = New ADODB.Recordset
       Set CRIBO = FLEXCONN.Execute("SELECT * FROM CRITBORR WHERE NCRITBO > 0 ORDER BY NCRITBO")
       CANCRI% = 0
       With CRIBO
          Do While Not .EOF
            CANCRI% = CANCRI% + 1
            ReDim Preserve STRI1$(CANCRI%), UBI1$(CANCRI%), STRI2$(CANCRI%), UBI2$(CANCRI%)
            '
            STRI1$(CANCRI%) = TRIM$(SAFETEXT$(!STRING1))
            UBI1$(CANCRI%) = UCase$(TRIM$(SAFETEXT$(!UBICA1)))
            STRI2$(CANCRI%) = "": UBI2$(CANCRI%) = ""
            If TRIM$(SAFETEXT$(!STRING2)) <> "" Then
               If TRIM$(SAFETEXT$(!UBICA2)) <> "" Then
                  STRI2$(CANCRI%) = TRIM$(SAFETEXT$(!STRING2))
                  UBI2$(CANCRI%) = UCase$(TRIM$(SAFETEXT$(!UBICA2)))
               End If
            End If
          .MoveNext
          Loop
       End With
       CRIBO.Close
       Set CRIBO = Nothing
       YACARLI% = 1
    End If
    '
    OKX% = 0                         ' DEVUELVE 1 SI LO PROCESO COMO MAIL SEGURO
    '
    CONDIX$ = "NUME_MAIL > 0 "
    CONDIX$ = CONDIX$ + " AND (PROS_DEST = '' OR PROS_DEST IS NULL) "
    CONDIX$ = CONDIX$ + " AND (BORR_MAIL = 0 OR BORR_MAIL IS NULL) "
    CONDIX$ = CONDIX$ + " ORDER BY NUME_MAIL ASC"
    FIN& = CantRegistros("BANPENPRO", CONDIX$)
    '
    Screen.MousePointer = 11
    Dim MAISELE As ADODB.Recordset
    Set MAISELE = New ADODB.Recordset
    Set MAISELE = FLEXCONN.Execute(FILTSQL$("SELECT * FROM BANPENPRO WITH(NOLOCK) WHERE " & CONDIX$))
    '
    With MAISELE
       JJ& = 0
       Screen.MousePointer = 11
       Do While Not .EOF
          JJ& = JJ& + 1
          Call TRACE(20, JJ&, FIN&)
          NUMAI& = XVALO(!NUME_MAIL)
          '
          LOBORRA% = 0
          ' LEE TODOS LOS DATOS
          ORENVI$ = SAFETEXT$(!ORIG_MAIL)
          DIELREM$ = SAFETEXT$(!DREL_MAIL)
          ASUNTO$ = SAFETEXT$(!ASUN_MAIL)
          TXTX$ = SAFETEXT$(!TEXT_MAIL)
          If TRIM$(DIELREM$) = "" Then    ' SI NO TIENE REMITENTE LO BORRA
              LOBORRA% = 1
              GoTo 39
          End If
          '
          For KKI% = 1 To CANCRI%
             TX1$ = "": TX2$ = ""
             If UBI1$(KKI%) = "ASUNTO" Then
                  TX1$ = ASUNTO$
                ElseIf UBI1$(KKI%) = "REMITENTE" Then
                  TX1$ = ORENVI$ + "<" + DIELREM$ + ">"
                ElseIf UBI1$(KKI%) = "TEXTO" Then
                  TX1$ = TXTX$
             End If
             If UBI2$(KKI%) = "ASUNTO" Then
                  TX2$ = ASUNTO$
                ElseIf UBI2$(KKI%) = "REMITENTE" Then
                  TX2$ = ORENVI$ + "<" + DIELREM$ + ">"
                ElseIf UBI2$(KKI%) = "TEXTO" Then
                  TX2$ = TXTX$
             End If
             '
             If InStr(UCase$(TX1$), UCase$(STRI1$(KKI%))) > 0 Then
                If STRI2$(KKI%) = "" Or InStr(UCase$(TX2$), UCase$(STRI2$(KKI%))) > 0 Then
                   LOBORRA% = 1
                   GoTo 39
                End If
             End If
             '
          Next KKI%
          '
39        If LOBORRA% > 0 Then
             CONDI$ = "NUME_MAIL = " & NUMAI&
             Call BORRAREGISTROS("BANPENPRO", CONDI$, RAFE&, "NOMESS")
             Call BORRAREGISTROS("ADJPENPRO", CONDI$, RAFE&, "NOMESS")
          End If
          '
       .MoveNext
       Loop
       '
    End With
    Screen.MousePointer = 1
    MAISELE.Close
    Set MAISELE = Nothing
    '
End Sub
'

