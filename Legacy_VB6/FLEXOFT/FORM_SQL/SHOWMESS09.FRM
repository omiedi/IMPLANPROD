VERSION 5.00
Begin VB.Form SHOWMESS09 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Muestra Mensaje Interno Recibido"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   8595
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
   ScaleHeight     =   7560
   ScaleWidth      =   8595
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00004000&
      Height          =   1800
      Left            =   7720
      ScaleHeight     =   1740
      ScaleWidth      =   945
      TabIndex        =   23
      Top             =   5985
      Width           =   1000
      Begin VB.CommandButton Command8 
         Caption         =   "1h"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   390
         Picture         =   "SHOWMESS09.frx":0000
         TabIndex        =   28
         ToolTipText     =   "Responder al Remitente"
         Top             =   120
         Width           =   340
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Mañana"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   90
         Picture         =   "SHOWMESS09.frx":014A
         TabIndex        =   27
         ToolTipText     =   "Responder al Remitente"
         Top             =   400
         Visible         =   0   'False
         Width           =   640
      End
      Begin VB.CommandButton Command2 
         Caption         =   "5'"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   90
         Picture         =   "SHOWMESS09.frx":0294
         TabIndex        =   26
         ToolTipText     =   "Responder al Remitente"
         Top             =   120
         Width           =   300
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   90
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   24
         Top             =   750
         Width           =   640
         Begin VB.Frame Frame4 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   0
            TabIndex        =   25
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -315
         Picture         =   "SHOWMESS09.frx":03DE
         Top             =   1020
         Width           =   1515
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Re-Enviar como E-Mail"
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
      Left            =   4725
      TabIndex        =   17
      ToolTipText     =   "Re-Enviar como Correo Electronico"
      Top             =   1470
      Width           =   2955
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004000&
      Height          =   1800
      Left            =   7720
      ScaleHeight     =   1740
      ScaleWidth      =   945
      TabIndex        =   9
      Top             =   0
      Width           =   1000
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
         Height          =   680
         Left            =   105
         Picture         =   "SHOWMESS09.frx":2300
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Cierra y Abandona la Aplicación"
         Top             =   150
         Width           =   640
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   90
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   10
         Top             =   1000
         Width           =   640
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   0
            TabIndex        =   11
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "SHOWMESS09.frx":244A
         Top             =   1340
         Width           =   1515
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
      Height          =   1580
      Left            =   105
      TabIndex        =   6
      Top             =   5880
      Width           =   7500
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
         Height          =   1110
         ItemData        =   "SHOWMESS09.frx":436C
         Left            =   105
         List            =   "SHOWMESS09.frx":436E
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   330
         Width           =   7245
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Acciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   4680
      TabIndex        =   4
      Top             =   105
      Width           =   2970
      Begin VB.CommandButton Command5 
         Caption         =   "Rep.All"
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
         Left            =   1025
         Picture         =   "SHOWMESS09.frx":4370
         TabIndex        =   21
         ToolTipText     =   "Responder a Todos"
         Top             =   315
         Width           =   900
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Reply"
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
         Left            =   105
         Picture         =   "SHOWMESS09.frx":44BA
         TabIndex        =   20
         ToolTipText     =   "Responder al Remitente"
         Top             =   315
         Width           =   900
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Resend"
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
         Left            =   1930
         Picture         =   "SHOWMESS09.frx":4604
         TabIndex        =   19
         ToolTipText     =   "Re-Enviar a Otro Usuario"
         Top             =   315
         Width           =   900
      End
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
         Left            =   105
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   840
         Width           =   2760
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Remitente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   105
      TabIndex        =   2
      Top             =   105
      Width           =   4470
      Begin VB.CommandButton Command10 
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   3820
         Picture         =   "SHOWMESS09.frx":474E
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Envia Confirmacion de Lectura al Remitente"
         Top             =   840
         Width           =   525
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
         Left            =   945
         TabIndex        =   16
         Top             =   330
         Width           =   3400
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
         Left            =   945
         TabIndex        =   15
         Top             =   900
         Width           =   2800
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
         Left            =   945
         TabIndex        =   14
         Top             =   1240
         Width           =   2800
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "C/Copia:"
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
         Top             =   945
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
         Left            =   105
         TabIndex        =   3
         Top             =   405
         Width           =   750
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Detalle de Mensaje a Remitir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3930
      Left            =   105
      TabIndex        =   0
      Top             =   1880
      Width           =   8340
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
         Height          =   3510
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   315
         Width           =   8100
      End
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
      Left            =   2940
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   645
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
      TabIndex        =   5
      Top             =   1890
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "SHOWMESS09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ENVIAINMEDIATO%

Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Sub ADJUNARCH(NROMAI&, FINAME$, RUTARCH$)
   XXX$ = LOADFILE$(RUTARCH$ + "\" + FINAME$)
   If XXX$ <> "" Then
10    NUMAPRO& = XVALO(RANDSTRING$(-6))
      If CantRegistros("ADJUNMAI", "NUME_MAIL = " & NUMAPRO&, "NOMESS") > 0 Then GoTo 10
      Call CreaRegistro("ADJUNMAI", "NUME_MAIL", NUMAPRO&, "NOMESS")
      '
      CONDI$ = "NUME_MAIL = " & NROMAI&
      NORDE& = 1 + XVALO(VALOBADA("ADJUNMAI", "MAX(NORD_MAIL)", CONDI$, "NOMESS"))
      Call ACTUREGISTRO("ADJUNMAI", "NORD_MAIL", "NUME_MAIL = " & NUMAPRO&, NORDE&, RAFE&, "NOMESS")
      Call ACTUREGISTRO("ADJUNMAI", "FILE_NAME", "NUME_MAIL = " & NUMAPRO&, FINAME$, RAFE, "NOMESS")
      Call ACTUREGISTRO("ADJUNMAI", "FILE_DATA", "NUME_MAIL = " & NUMAPRO&, XXX$, RAFE&, "NOMESS")
      Call ACTUREGISTRO("ADJUNMAI", "NUME_MAIL", "NUME_MAIL = " & NUMAPRO&, NROMAI&, RAFE&, "NOMESS")
   End If
   '
End Sub
'
Sub CIERRAPLI(TAPLI$)
    '
    VUELTA% = 0
    '
10  On Error Resume Next
    AppActivate TAPLI$
    If Err < 1 Then
       On Error GoTo 0
       SendKeys "%{F4}", True
       VUELTA% = VUELTA% + 1
       If VUELTA% < 3 Then GoTo 10
    End If
    On Error GoTo 0
    '
End Sub

Private Sub Command1_Click()
   Command1.Enabled = False
   PREPAMAI09.Show 1
   If PREPAMAI09.Label1 = "OK" Then
      PERSODES$ = PREPAMAI09.Combo1.TEXT
      FISECOM09.ORIMAI = "SHOWMESS09"
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

Private Sub Command10_Click()
   '
   Command10.Enabled = False
   '
   USNU% = USNBR% Mod 100
   NUMENSAJ& = XVALO(Label13)
   CONDI$ = "NUME_MENSA = " & NUMENSAJ&
   CONDI$ = CONDI$ + " AND PARA_MENSA = " & USNU%
   Call ACTUREGISTRO("MENSAINT", "CONF_RECEP", CONDI$, AHORA$, RAFE&, "NOMESS")
   '
   CONDI$ = "NUME_MENSA = " & NUMENSAJ&
   CONDI$ = CONDI$ + " AND CCOP_MENSA = " & USNU%
   Call ACTUREGISTRO("MENSAINT", "CONF_RECOP", CONDI$, AHORA$, RAFE&, "NOMESS")
   '
   Command10.Enabled = True
   '
End Sub

Sub ENVIAMENSAJE(TMENSA$, ASUNTO$, DESTI1%, DESTI2%, DESTI3%, SOLICO%, FHORPROG$, Optional LISTADEADJUN$)
   '
   Call CREACAMPO("", "MENSAINT", "NUME_MENSA", 4, 4, 1)
   Call CREACAMPO("", "MENSAINT", "ASUN_MENSA", 10, 256)
   Call CREACAMPO("", "MENSAINT", "PARA_MENSA", 3, 2, 1)
   Call CREACAMPO("", "MENSAINT", "CCOP_MENSA", 3, 2, 1)
   Call CREACAMPO("", "MENSAINT", "TEXT_MENSA", 12, 0)
   Call CREACAMPO("", "MENSAINT", "SOLI_CONF", 3, 2)
   Call CREACAMPO("", "MENSAINT", "CONF_RECEP", 8, 0, 1)       ' FECHA Y HORA CONFIRMACION DE RECEPCION
   Call CREACAMPO("", "MENSAINT", "CONF_RECOP", 8, 0, 1)       ' FECHA Y HORA CONFIRMACION DE RECEPCION COPIA
   Call CREACAMPO("", "MENSAINT", "FHOR_PREPA", 8, 0, 1)       ' FECHA Y HORA QUE FUE PREPARADO
   Call CREACAMPO("", "MENSAINT", "USUA_PREPA", 3, 2, 1)       ' USUARIO QUE LO PREPARÓ
   Call CREACAMPO("", "MENSAINT", "FHOR_SOLENVI", 8, 0, 1)     ' FECHA Y HORA PROXIMO ENVIO PROGRAMADO
   Call CREACAMPO("", "MENSAINT", "FHOR_ULTENVI", 8, 0, 0)     ' FECHA Y HORA ULTIMO PRESENTACION
   Call CREACAMPO("", "MENSAINT", "FHOR_ULTENCO", 8, 0, 0)     ' FECHA Y HORA ULTIMO PRESENTACION
   '
   Call CREACAMPO("", "ADJUNMEN", "NUME_MENSA", 4, 4, 1)
   Call CREACAMPO("", "ADJUNMAI", "NORD_MENSA", 3, 2, 1)
   Call CREACAMPO("", "ADJUNMAI", "FILE_NAME", 10, 256, 0)
   '
   NUPROMEN& = 1 + XVALO(VALOBADA("MENSAINT", "MAX(NUME_MENSA)", "NUME_MENSA > 0", "NOMESS"))
   CONDI$ = "NUME_MAIL = " & NUPROMEN&
   Call CreaRegistro("BANDENVI", "NUME_MENSA", NUPROMEN&, "NOMESS")
   ' Label10 = TRIM$(Str$(NUPROMAI&))
   '
   Call ACTUREGISTRO("MENSAINT", "NUME_MENSA", CONDI$, NUPROMEN&, RAFE&, "NOMESS")
   Call ACTUREGISTRO("MENSAINT", "ASUN_MENSA", CONDI$, Left$(ASUNTO$, 256), RAFE&, "NOMESS")
   Call ACTUREGISTRO("MENSAINT", "PARA_MENSA", CONDI$, DESTI1%, RAFE&, "NOMESS")
   Call ACTUREGISTRO("MENSAINT", "CCOP_MENSA", CONDI$, DESTI2%, RAFE&, "NOMESS")
   Call ACTUREGISTRO("MENSAINT", "TEXT_MENSA", CONDI$, TMENSA$, RAFE&, "NOMESS")
   Call ACTUREGISTRO("MENSAINT", "SOLI_CONF", CONDI$, SOLICO%, RAFE&, "NOMESS")
   Call ACTUREGISTRO("MENSAINT", "CONF_RECEP", CONDI$, FETEXCOR1$(FECHANUM("01/01/80")), RAFE&, "NOMESS")
   Call ACTUREGISTRO("MENSAINT", "FHOR_PREPA", CONDI$, AHORA$, RAFE&, "NOMESS")           ' F-HOR PREPARADO
   Call ACTUREGISTRO("MENSAINT", "USUA_PREPA", CONDI$, USNBR% Mod 100, RAFE&, "NOMESS")   ' USUARIO PREPARÓ
   Call ACTUREGISTRO("MENSAINT", "FHOR_SOLENVI", CONDI$, FHORPROG$, RAFE&, "NOMESS")       ' PROXIMO ENVIO
   Call ACTUREGISTRO("MENSAINT", "FHOR_ULTENVI", CONDI$, NUNCA$, RAFE&, "NOMESS")
   Call ACTUREGISTRO("MENSAINT", "FHOR_ULTENCO", CONDI$, NUNCA$, RAFE&, "NOMESS")
   '

End Sub

Private Sub Command2_Click()
   '
   NUMENSAJ& = XVALO(Label13)
   FHORPROG = Now + 5 / (24 * 60)
   FPRX$ = Format(FHORPROG, "DD/MM/YY  HH:MM:SS")
   FEXI% = FECHANUM(Left$(FPRX$, 8))
   NUFHOR = FETEXCOR1$(FEXI%) + Mid$(FPRX$, 9)
   '
   CONDI$ = "NUME_MENSA = " & NUMENSAJ&
   Call ACTUREGISTRO("MENSAINT", "FHOR_SOLENVI", CONDI$, NUFHOR, RAFE&, "NOMESS")
   Call ACTUREGISTRO("MENSAINT", "ENVIA_YA", CONDI$, 1, RAFE&, "NOMESS")
   '
   Unload Me
   
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   If Combo1.TEXT <> "" Then
        TMENSA$ = Text2
        If TRIM$(TMENSA$) = "" Then Call MENSERR(24, "Falta Texto de Mensaje."): GoTo 99
        DESTI1% = XVALO(Left$(Combo1.TEXT, 3))
        DESTIX$ = USERNAM$(DESTI1%)
        DESTI2% = 0:   DESTI3% = 0
        SOLICO% = 1
        FHORPROG$ = AHORA$
        '
        LISTADEADJUN$ = ""
        For KKI% = 1 To List3.ListCount
           If LISTADEADJUN$ <> "" Then
              LISTADEADJUN$ = LISTADEADJUN$ + ";"
           End If
           LISTADEADJUN$ = LISTADEADJUN$ + List3.List(KKI% - 1)
        Next KKI%
        '
        Call ENVIAMENSAJE(TMENSA$, ASUNTO$, DESTI1%, DESTI2%, DESTI3%, SOLICO%, FHORPROG$, LISTADEADJUN$)
        Call COMUNI("Mensaje Re-Enviado a '" + DESTIX$ + "'.")
      Else
        Call MENSERR(24, "Debe Indicar Destinatario del Re-Envío")
   End If
99 Command3.Enabled = True
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   Call Command10_Click      ' pone marca de confirmacion de lectura
   On Error Resume Next
   MENSAIN09.Combo1.TEXT = Label14.Caption
   On Error GoTo 0
   '
   MENSAIN09.Label2.Font = MENSAIN09.Text2.Font
   MENSAIN09.Label2.FontSize = MENSAIN09.Text2.FontSize
   MENSAIN09.Label2.FontBold = MENSAIN09.Text2.FontBold
   MENSAIN09.Label2 = ""
   While MENSAIN09.Label2.Width < MENSAIN09.Text2.Width - 500
      MENSAIN09.Label2.Caption = MENSAIN09.Label2.Caption + "="
   Wend
   MENSAIN09.Text2 = CRLF + CRLF + CRLF + MENSAIN09.Label2.Caption + CRLF + Text2
   MENSAIN09.Show 1
   Unload Me
   Command4.Enabled = True
End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
   On Error Resume Next
   MENSAIN09.Combo2.TEXT = Label16.Caption
   MENSAIN09.Combo3.TEXT = Label17.Caption
   On Error GoTo 0
   Call Command4_Click
   Command5.Enabled = True
End Sub

Private Sub Command6_Click()
   '
   NUMENSAJ& = XVALO(Label13)
   FHORPROG = 1 + Int(Now)
   FPRX$ = Format(FHORPROG, "DD/MM/YY  HH:MM:SS")
   FEXI% = FECHANUM(Left$(FPRX$, 8))
   NUFHOR = FETEXCOR1$(FEXI%) + Mid$(FPRX$, 9)
   '
   CONDI$ = "NUME_MENSA = " & NUMENSAJ&
   Call ACTUREGISTRO("MENSAINT", "FHOR_SOLENVI", CONDI$, NUFHOR, RAFE&, "NOMESS")
   Call ACTUREGISTRO("MENSAINT", "ENVIA_YA", CONDI$, 1, RAFE&, "NOMESS")
   '
   Unload Me
   
End Sub

Private Sub Command7_Click()
   Call CIERRARCH("*.*")
   Hide
End Sub


Private Sub Command9_Click()
   '
   Call SELECHO("SELFECHA")
   VDEV$ = VALACT1$("SELFECHA")
   If VDEV$ <> "" Then
       Label4.Caption = VDEV$
   End If

End Sub

Private Sub Command8_Click()
   '
   NUMENSAJ& = XVALO(Label13)
   FHORPROG = Now + 60 / (24 * 60)
   FPRX$ = Format(FHORPROG, "DD/MM/YY  HH:MM:SS")
   FEXI% = FECHANUM(Left$(FPRX$, 8))
   NUFHOR = FETEXCOR1$(FEXI%) + Mid$(FPRX$, 9)
   '
   CONDI$ = "NUME_MENSA = " & NUMENSAJ&
   Call ACTUREGISTRO("MENSAINT", "FHOR_SOLENVI", CONDI$, NUFHOR, RAFE&, "NOMESS")
   Call ACTUREGISTRO("MENSAINT", "ENVIA_YA", CONDI$, 1, RAFE&, "NOMESS")
   '
   Unload Me
   
End Sub

Private Sub Form_Activate()
   '
   If ATENTI% > 0 Then
     Left = 0: Top = 0
     Call CENTRAFORM(Me)
     ATENTI% = 0
   End If
   '
   VEMOFREN% = 0
20 If VEMOFREN% < 10 Then          ' que lo haga 10 veces
      VEMOFREN% = VEMOFREN% + 1
      On Error Resume Next
      Me.ZOrder (0)
      AppActivate Caption
      SendKeys Chr$(0)
      On Error GoTo 0
      OPENFORMS = DoEvents
      GoTo 20
   End If
   '
End Sub

Private Sub Form_Load()
   '
    Call CENTRAFORM(Me)
    Call CARLIUS(Combo1)
    '
    If USNBR% Mod 100 = 1 Then
       Command6.Visible = True
    End If
End Sub
'
Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Cancel = 0
End Sub

Private Sub Label4_DblClick()
   Label4 = ""
   Text1 = ""
End Sub

Private Sub Label13_Change()
   '
   USACT% = USNBR% Mod 100
   NUMENSA& = XVALO(Label13)
   CONDI$ = "NUME_MENSA = " & NUMENSA&
   REMITE% = XVALO(VALOBADA("MENSAINT", "USUA_PREPA", CONDI$, "NOMESS")) ' USUARIO PREPARÓ
   Label14 = USERNUNAM$(REMITE%)
   '
   DESTI1% = XVALO(VALOBADA("MENSAINT", "PARA_MENSA", CONDI$, "NOMESS"))
   DESTI2% = XVALO(VALOBADA("MENSAINT", "CCOP_MENSA", CONDI$, "NOMESS"))
   If DESTI1% = USACT% Then
      DESTI1% = DESTI2%: DESTI2% = 0
   End If
   If DESTI2% = USACT% Then DESTI2% = 0
   Label16 = USERNUNAM$(DESTI1%)
   Label17 = USERNUNAM$(DESTI2%)
   '
   Text2 = VALOBADA("MENSAINT", "TEXT_MENSA", CONDI$, "NOMESS")
   '   Call ACTUREGISTRO("MENSAINT", "SOLI_CONF", CONDI$, SOLICO%, RAFE&, "NOMESS")
   '   Call ACTUREGISTRO("MENSAINT", "CONF_RECEP", CONDI$, FETEXCOR1$(FECHANUM("01/01/80")), RAFE&, "NOMESS")
   '   Call ACTUREGISTRO("MENSAINT", "FHOR_PREPA", CONDI$, AHORA$, RAFE&, "NOMESS")           ' F-HOR PREPARADO
   '   Call ACTUREGISTRO("MENSAINT", "FHOR_SOLENVI", CONDI$, FHORPROG$, RAFE&, "NOMESS")       ' PROXIMO ENVIO
   '   Call ACTUREGISTRO("MENSAINT", "FHOR_ULTENVI", CONDI$, NUNCA$, RAFE&, "NOMESS")
   '   Call ACTUREGISTRO("MENSAINT", "FHOR_ULTENCO", CONDI$, NUNCA$, RAFE&, "NOMESS")
   '
   List3.Clear
   LISTADJUN$ = VALOBADA("MENSAINT", "LISADJUN", CONDI$, "NOMESS")
10 PPXX% = InStr(LISTADJUN$, ";")
   If PPXX% < 1 Then PPXX% = 1 + Len(LISTADJUN$)
   FINAM$ = Left$(LISTADJUN$, PPXX% - 1)
   If FINAM$ <> "" Then
      List3.AddItem FINAM$
      LISTADJUN$ = Mid$(LISTADJUN$, PPXX% + 1)
      GoTo 10
   End If
   '


End Sub

Private Sub List3_DblClick()
   DOCUACT$ = TRIM$(Mid$(List3.TEXT, 129))
   Call MUESTRADOC4(DOCUACT$)
End Sub

Sub MUESTRADOC4(DOCUACT$)
    '
    Call ShellExecute(Me.hwnd, "Open", DOCUACT$, "", "", 1)
    '
End Sub

Private Sub Text2_GotFocus()
    Command7.SetFocus
End Sub
