VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form SHINMAIL09 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Muestra Mensaje de Correo Electrónico Recibido"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9375
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
   ScaleHeight     =   7890
   ScaleWidth      =   9375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   8600
      Picture         =   "SHINMAIL09.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Cierra y Abandona la Aplicación"
      Top             =   1360
      Width           =   645
   End
   Begin VB.Frame Frame5 
      Caption         =   "Acciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1060
      Left            =   3480
      TabIndex        =   26
      Top             =   5880
      Width           =   5805
      Begin VB.Frame Frame6 
         Caption         =   "Re-Asignar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1060
         Left            =   3240
         TabIndex        =   32
         Top             =   0
         Width           =   2550
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   600
            Width           =   2280
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Ok"
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
            Left            =   1850
            Picture         =   "SHINMAIL09.frx":066A
            TabIndex        =   33
            ToolTipText     =   "Re-Enviar a Otro Usuario"
            Top             =   180
            Width           =   540
         End
      End
      Begin VB.CommandButton Command18 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   2520
         Picture         =   "SHINMAIL09.frx":07B4
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Marca de No Responder"
         Top             =   360
         Width           =   550
      End
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
         Height          =   570
         Left            =   1320
         Picture         =   "SHINMAIL09.frx":0ABE
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Responder a Todos"
         Top             =   360
         Width           =   550
      End
      Begin VB.CommandButton Command2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   720
         Picture         =   "SHINMAIL09.frx":0DC8
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Responder al Remitente"
         Top             =   360
         Width           =   550
      End
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
         Height          =   570
         Left            =   1920
         Picture         =   "SHINMAIL09.frx":0F12
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Re-Enviar a Otgro Destino"
         Top             =   360
         Width           =   550
      End
      Begin VB.CommandButton Command3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   120
         Picture         =   "SHINMAIL09.frx":1354
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Confirmar Lectura"
         Top             =   360
         Width           =   550
      End
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   8620
      ScaleHeight     =   75
      ScaleWidth      =   585
      TabIndex        =   17
      Top             =   7080
      Width           =   650
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Recibido en "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   3480
      TabIndex        =   14
      Top             =   7000
      Width           =   3300
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Fech-Hora:"
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
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   1035
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
         Height          =   255
         Left            =   1200
         TabIndex        =   15
         Top             =   315
         Width           =   1935
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Archivos Adjuntos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1880
      Left            =   120
      TabIndex        =   7
      Top             =   5880
      Width           =   3250
      Begin VB.ListBox List3 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1420
         IntegralHeight  =   0   'False
         Left            =   105
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   315
         Width           =   3040
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
      Height          =   525
      Left            =   8600
      Picture         =   "SHINMAIL09.frx":165E
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Cierra y Abandona la Aplicación"
      Top             =   210
      Width           =   645
   End
   Begin VB.Frame Frame3 
      Caption         =   "Confirmacion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   6840
      TabIndex        =   5
      Top             =   7000
      Width           =   1610
      Begin VB.CheckBox Check1 
         Caption         =   "Solicitada"
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
         Left            =   240
         TabIndex        =   11
         Top             =   280
         Width           =   1275
      End
   End
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
      TabIndex        =   1
      Top             =   105
      Width           =   8350
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
         TabIndex        =   23
         Top             =   960
         Width           =   7005
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
         TabIndex        =   22
         Top             =   630
         Width           =   7005
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
         TabIndex        =   21
         Top             =   1270
         Width           =   7005
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
         TabIndex        =   20
         Top             =   300
         Width           =   5205
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
         Left            =   6960
         TabIndex        =   13
         Top             =   300
         Width           =   1245
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
         Left            =   6360
         TabIndex        =   12
         Top             =   360
         Width           =   555
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
         TabIndex        =   10
         Top             =   1030
         Width           =   750
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
         TabIndex        =   4
         Top             =   1365
         Width           =   960
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
         TabIndex        =   3
         Top             =   405
         Width           =   855
      End
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
         TabIndex        =   2
         Top             =   720
         Width           =   750
      End
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
      Height          =   3930
      Left            =   105
      TabIndex        =   0
      Top             =   1850
      Width           =   9180
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   3450
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Visible         =   0   'False
         Width           =   8895
         ExtentX         =   15690
         ExtentY         =   6085
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
         Height          =   3450
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         Text            =   "SHINMAIL09.frx":17A8
         Top             =   360
         Width           =   8895
      End
   End
   Begin VB.Label Label139 
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
      Left            =   8610
      TabIndex        =   35
      Top             =   750
      Width           =   645
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
      Left            =   8610
      TabIndex        =   19
      Top             =   1050
      Width           =   645
   End
   Begin VB.Image Image4 
      Height          =   390
      Left            =   8180
      Picture         =   "SHINMAIL09.frx":17AE
      Top             =   7350
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
      TabIndex        =   6
      Top             =   1890
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "SHINMAIL09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RESTODELMAIL$
Dim ENVIOMASIVO%

Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Sub ENVIAMAIL()
      
   MAI_DESTI$ = TRIM$(Label14)
   MAI_DIREL$ = TRIM$(Label16)
   MAI_COPIA$ = TRIM$(Label17)
   MAI_ASUNT$ = TRIM$(Label15)
   Call FRATEXTO("." & Text2, 88, 1, 1)
   TTXX$ = ""
   For KKI% = 1 To CARETEX%
     TTXX$ = TTXX$ + RETEX$(KKI%) + CRLF$
   Next KKI%
   TTYY$ = REPLACAR$(TTXX$, Chr$(13), "~")
   TTYY$ = REPLACAR$(TTXX$, Chr$(10), "")
   TTYY$ = REPLACAR$(TTXX$, "(", "-")
   TTYY$ = REPLACAR$(TTXX$, ")", "-")
   MAI_TEXTO$ = TTYY$
   '
   Call GENERAMAIL
   '
End Sub

Private Sub Command16_Click()
    Call SELECHO("SELFECHA")
    VDEV$ = VALACT1$("SELFECHA")
    If VDEV$ <> "" Then
        Text7.TEXT = VDEV$
    End If
End Sub

Private Sub Command1_Click()
   Command1.Enabled = False
   PREPAMAI09.Show 1
   If PREPAMAI09.Label1 = "OK" Then
      PERSODES$ = PREPAMAI09.Combo1.TEXT
      FISECOM09.ORIMAI = "SHINMAIL09"
      For KKI% = 1 To FISECOM09.Text6.Count
         If TRIM$(FISECOM09.Text6(KKI% - 1)) <> "" Then
            If FISECOM09.Text6(KKI% - 1) = PREPAMAI09.Combo1.TEXT Then
               Call FISECOM09.Text6_DblClick(KKI% - 1)
               GoTo 80
            End If
         End If
      Next KKI%
      Call FISECOM09.Text5_DblClick
      '
80    PREPAMAI09.Label1 = ""
   End If
   Command1.Enabled = True
End Sub

Private Sub Command18_Click()
    If DERACCE%("NORESPON", "Marcar Mail con 'No Responder'") >= 2 Then
       ANUMAIL09.Label13 = Label13
       ANUMAIL09.Label26 = Label26
       ANUMAIL09.Left = Me.Left + (Me.Width - ANUMAIL09.Width) / 2
       ANUMAIL09.Top = Me.Top + Me.Height - ANUMAIL09.Height - 30
       On Error Resume Next
       ANUMAIL09.Combo1 = Label14
       On Error GoTo 0
       ANUMAIL09.Show 1
       If ANUMAIL09.Label13 = "ANUL" Then
          Unload ANUMAIL09
          Hide
       End If
    End If
End Sub

Private Sub Command2_Click()
    '
    Label139 = ""
    FISECOM09.Label26 = Label26
    CODCLI$ = Label26
    DESTI$ = Label14
    DIMAI$ = Label16
    ASUNTO$ = "RE: " + Label15
    ORIMESS$ = Text1.TEXT
    Call FISECOM09.CREAMAIL(CODCLI$, DESTI$, DIMAI$, ASUNTO$, ORIMESS$)
    '
    MAILMESS07.Show 1
    If XVALO(Label139) > 0 Then
       Hide
    End If
    '
End Sub

Private Sub Command4_Click()
AA1 = Text1
End Sub

Private Sub Command5_Click()
   Call STDFORM("PRINMAIL", Label13)
End Sub
'
Private Sub Command7_Click()
   Hide
End Sub

Private Sub Command9_Click()
   '
   Call SELECHO("SELFECHA")
   VDEV$ = VALACT1$("SELFECHA")
   If VDEV$ <> "" Then
       Label4.Caption = VDEV$
   End If
   '
End Sub

Private Sub Command8_Click()
   Command3.Enabled = False
   If Combo1.TEXT <> "" Then
        TMENSA$ = Text1
        If TRIM$(TMENSA$) = "" Then Call MENSERR(24, "Falta Texto de Mensaje."): GoTo 99
        '
        DESTI1% = XVALO(Left$(Combo1.TEXT, 3))
        CONDI$ = "NUME_MAIL = " & Label13
        Call ACTUREGISTRO("BANDENTRA", "USUA_DESTI", CONDI$, DESTI1%, RAFE&)
        DESTIX$ = USERNAM$(DESTI1%)
        DESTI2% = 0:   DESTI3% = 0
        SOLICO% = 1
        FHORPROG$ = AHORA$
        '
        TMENSA$ = DESTIX$ + ":" + CRLF + CRLF
        TMENSA$ = TMENSA$ + "Te he re-asignado el mail recibido el " + TRIM$(Label2) + " de " + Label14 + "<" + Label16 + "> - Asunto: " + Label15 + " para su procesamiento y respuesta." + CRLF + CRLF
        TMENSA$ = TMENSA$ + "Cualquier duda, consultame." + CRLF + CRLF
        TMENSA$ = TMENSA$ + "Muchas gracias."
        '
        ASUNTO$ = "Re-Asignación de Mail de " + Label14
        '
        Call SHOWMESS09.ENVIAMENSAJE(TMENSA$, ASUNTO$, DESTI1%, DESTI2%, DESTI3%, SOLICO%, FHORPROG$, LISTADEADJUN$)
        Call COMUNI("Mail Re-Asignado a '" + DESTIX$ + "'.")
      Else
        Call MENSERR(24, "Debe Indicar Destinatario del Re-Envío")
   End If
99 Command3.Enabled = True
End Sub

Private Sub Form_Activate()
   If ATENTI% > 0 Then
     Left = 0
     Top = 0
     'Call CENTRAFORM(Me)
     ATENTI% = 0
     Text2.SetFocus
   End If
End Sub

Private Sub Form_Load()
   '
   Call CARLIUS(Combo1)
   '
End Sub
'
Private Sub Form_Unload(Cancel As Integer)
'   Call CIERRARCH("*.*")
'   Call BAJALDISCO
   Cancel = 0
End Sub

Private Sub Label13_Change()
    WebBrowser1.Visible = False
    NUMAI& = XVALO(Label13)
    CONDI$ = "NUME_MAIL = " & NUMAI&
    If CantRegistros("BANDENTRA", CONDI$, "NOMESS") > 0 Then
          Caption = "Muestra Mensaje de Correo Electronico #" & TRIM$(Str$(NUMAI&))
          Label26 = TRIM$(VALOBADA("BANDENTRA", "Codi_Cli", CONDI$, "NOMESS"))
          Label17 = TRIM$(VALOBADA("BANDENTRA", "PARA_MAIL", CONDI$, "NOMESS"))
          Label15 = TRIM$(VALOBADA("BANDENTRA", "ASUN_MAIL", CONDI$, "NOMESS"))
          Label16 = TRIM$(VALOBADA("BANDENTRA", "DREL_MAIL", CONDI$, "NOMESS"))
          Label14 = TRIM$(VALOBADA("BANDENTRA", "ORIG_MAIL", CONDI$, "NOMESS"))
          Label2 = FORMATNUM(VALOBADA("BANDENTRA", "FHOR_RECEP", CONDI$, "NOMESS"), "D17")
          TXMAIL$ = VALOBADA("BANDENTRA", "TEXT_MAIL", CONDI$, "NOMESS")
          HTMAIL$ = VALOBADA("BANDENTRA", "HTMT_MAIL", CONDI$, "NOMESS")
          '
          If TRIM$(HTMAIL$) <> "" And UCase$(Left$(TRIM$(HTMAIL$), 5)) <> "TEXT2" Then
               Text1.TEXT = TXMAIL$
               RUTAR$ = "C:\FLEXOFT\ARADJUN\"
               Call SAVEFILE(RUTAR$ + "HTMMAIL.HTM", HTMAIL$) ' GUARDA EL PRINCIPAL
               Call CARGALISTA(List3, "ADJUENTRA", "NORD_MAIL/F3; FILE_NAME/A128", CONDI$, "NOMESS")
               For KKI% = 1 To List3.ListCount
                  NORMAI& = XVALO(Left$(List3.List(KKI% - 1), 3))
                  CONDI2$ = "NUME_MAIL = " & NUMAI& & " AND NORD_MAIL = " & NORMAI&
                  FINAM$ = VALOBADA("ADJUENTRA", "FILE_NAME", CONDI2$, "NOMESS")
                  XXXX$ = VALOBADA("ADJUENTRA", "FILE_DATA", CONDI2$, "NOMESS")
                  If FINAM$ <> "" Then
                     If XXXX$ <> "" Then
                        Call SAVEFILE(RUTAR$ + FINAM$, XXXX$)
                     End If
                  End If
               Next KKI%
               WebBrowser1.Navigate "C:\FLEXOFT\ARADJUN\HTMMAIL.HTM"
               WebBrowser1.Visible = True
               
            Else
               Text1.TEXT = TXMAIL$
               Text1.Refresh
          End If
          '
          Check1.Value = XVALO(VALOBADA("BANDENTRA", "SOLI_CONF", CONDI$, "NOMESS"))
          Call CARGALISTA(List3, "ADJUENTRA", "NORD_MAIL/F3; FILE_NAME/A128", CONDI$, "NOMESS")
          '
       Else
          Call MENSERR(24, "Imposible Abrir Mensaje Enviado Numero '" & TRIM$(Str$(NUMAI&)) & "'")
   End If
End Sub

Private Sub Label4_DblClick()
   Label4 = ""
   Text1 = ""
End Sub


Private Sub Label26_DblClick()
   If TRIM$(Label26) <> "" Then
      ANOTASEG09.Label26 = Label26
      ANOTASEG09.Show 1
   End If
End Sub

Private Sub List3_DblClick()
   On Error Resume Next
   MkDir LUCOM$ + "ARADJUN"
   On Error GoTo 0
   RUTAR$ = LUCOM$ + "ARADJUN\"
   '
   NUMAI& = XVALO(Label13)
   NORMAI& = XVALO(Left$(List3.TEXT, 3))
   CONDI$ = "NUME_MAIL = " & NUMAI& & " AND NORD_MAIL = " & NORMAI&
   FINAM$ = VALOBADA("ADJUENTRA", "FILE_NAME", CONDI$, "NOMESS")
   XXXX$ = VALOBADA("ADJUENTRA", "FILE_DATA", CONDI$, "NOMESS")
   If FINAM$ <> "" Then
      If XXXX$ <> "" Then
         Call SAVEFILE(RUTAR$ + FINAM$, XXXX$)
         DOCUACT$ = RUTAR$ + FINAM$
         Call MUESTRADOC4(DOCUACT$)
      End If
   End If
   '
End Sub

Sub MUESTRADOC4(DOCUACT$)
    '
    Call ShellExecute(Me.hwnd, "Open", DOCUACT$, "", "", 1)
    '
End Sub

Private Sub Text12_DblClick()
    Call SELECHO("SELFECHA")
    VDEV$ = VALACT1$("SELFECHA")
    If VDEV$ <> "" Then
        Text12.TEXT = VDEV$
    End If
End Sub

Sub ARMADESTI(CODICLI$, Optional DESTI$)
   '
   KOKLI$ = TRIM$(CODICLI$)
   CONDICION$ = "CODI_CLI = '" & KOKLI$ & "'"
   '
   EMPRE$ = TRIM$(VALOBADA("BSEGCOM", "RASO_CLI", CONDICION$)) 'TRIM$(Mid$(RELE$, 3, 59))
   CALNUM$ = TRIM$(VALOBADA("BSEGCOM", "DOMI_CLI", CONDICION$)) 'TRIM$(Mid$(RELE$, 129, 48))
   CPOST$ = TRIM$(VALOBADA("BSEGCOM", "Cpos_Cli", CONDICION$)) 'TRIM$(Mid$(RELE$, 177, 8))
   If CPOST$ <> "" Then CPOST$ = CPOST$ + " "
   LOCAX$ = CPOST$ + TRIM$(VALOBADA("BSEGCOM", "Loca_Cli", CONDICION$)) ' CPOST$ + TRIM$(Mid$(RELE$, 185, 32))
   PROPA$ = TRIM$(VALOBADA("BSEGCOM", "Prov_Cli", CONDICION$)) 'TRIM$(Mid$(RELE$, 217, 32))
   PRO$ = TRIM$(ECOARCH$("provinc", PROPA$))
   If PRO$ <> "" Then PRO$ = PRO$ + " - "
   
   PRO$ = PRO$ + TRIM$(VALOBADA("BSEGCOM", "Pais_Cli", CONDICION$)) 'TRIM$(PROPA$ + TRIM$(Mid$(RELE$, 249, 24)))
   '
   TTXX$ = ""
   If EMPRE$ <> "" Then TTXX$ = TTXX$ + EMPRE$ + Chr$(13) + Chr$(10)
   If CALNUM$ <> "" Then TTXX$ = TTXX$ + CALNUM$ + Chr$(13) + Chr$(10)
   If LOCAX$ <> "" Then TTXX$ = TTXX$ + LOCAX$ + Chr$(13) + Chr$(10)
   If PRO$ <> "" Then TTXX$ = TTXX$ + PRO$ + Chr$(13) + Chr$(10)
   '
   ASUNTX$ = TRIM$(ASUNTO$)
   If ASUNTX$ = "" Then ASUNTX$ = CONTROL$("SEGUICOM", "ASUNTO")
   If ASUNTX$ = "" Then ASUNTX$ = "FLEXOFT"
   If ASUNTX$ = "FLEXOFT" Then
      If TRIM$(DESTI$) <> "" Then
         ASUNTX$ = ASUNTX$ + " para " + DESTI$
      End If
   End If
   '
   SHOWMAIL09.Label15 = ASUNTX$
   '
   HOII% = HOY(HOS$)
   FEC = FETEXTO(HOII%)
   LUGARYFECH$ = AJUSTD("Buenos Aires - " + FEC, 100)
   TTYY$ = LUGARYFECH$ + vbCrLf
   '
   TTYY$ = TTYY$ + "Sr./Sra." + Chr$(13) + Chr$(10)
   TTYY$ = TTYY$ + DESTI$ + Chr$(13) + Chr$(10)
   '
   TTYY$ = TTYY$ + TTXX$ + RESTODELMAIL$
   Text2 = TTYY$
   '
End Sub


