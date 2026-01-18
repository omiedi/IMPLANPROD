VERSION 5.00
Begin VB.Form MENSAIN09 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generar Mensaje Interno"
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
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004000&
      Height          =   1800
      Left            =   7720
      ScaleHeight     =   1740
      ScaleWidth      =   945
      TabIndex        =   20
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
         Picture         =   "MENSAIN09.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   23
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
         TabIndex        =   21
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
            TabIndex        =   22
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "MENSAIN09.frx":014A
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
      Width           =   8400
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
         ItemData        =   "MENSAIN09.frx":206C
         Left            =   105
         List            =   "MENSAIN09.frx":206E
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   330
         Width           =   7035
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Agregar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   520
         Left            =   7245
         TabIndex        =   8
         Top             =   330
         Width           =   1000
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Quitar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   520
         Left            =   7245
         TabIndex        =   7
         Top             =   900
         Width           =   1000
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Enviar"
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
      Left            =   5040
      TabIndex        =   4
      Top             =   105
      Width           =   2550
      Begin VB.CheckBox Check2 
         Caption         =   "Inmediato"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   225
         TabIndex        =   26
         Top             =   240
         Width           =   1350
      End
      Begin VB.Frame Frame10 
         Caption         =   "Programar"
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
         Left            =   140
         TabIndex        =   15
         Top             =   900
         Width           =   2265
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
            Height          =   250
            Left            =   1365
            TabIndex        =   17
            Top             =   315
            Width           =   465
         End
         Begin VB.CommandButton Command9 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   260
            Left            =   1080
            TabIndex        =   16
            ToolTipText     =   "Lista Completa de Contactos"
            Top             =   315
            Width           =   150
         End
         Begin VB.Label Label1 
            Caption         =   "Hs"
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
            Left            =   1890
            TabIndex        =   19
            Top             =   315
            Width           =   225
         End
         Begin VB.Label Label4 
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
            Left            =   105
            TabIndex        =   18
            Top             =   315
            Width           =   990
         End
      End
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
         Height          =   705
         Left            =   1680
         Picture         =   "MENSAIN09.frx":2070
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Guardar y Enviar Mensaje"
         Top             =   210
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Confirmar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   225
         TabIndex        =   11
         Top             =   500
         Value           =   1  'Checked
         Width           =   1230
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Destinatario"
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
      Width           =   4785
      Begin VB.ComboBox Combo3 
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
         Left            =   1155
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   1220
         Width           =   3500
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
         Left            =   1155
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   325
         Width           =   3500
      End
      Begin VB.ComboBox Combo2 
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
         Left            =   1155
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   855
         Width           =   3500
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
         Left            =   315
         TabIndex        =   10
         Top             =   930
         Width           =   750
      End
      Begin VB.Label Label3 
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
         Left            =   210
         TabIndex        =   3
         Top             =   405
         Width           =   855
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Height          =   150
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Visible         =   0   'False
      Width           =   30
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
Attribute VB_Name = "MENSAIN09"
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

Private Sub Command10_Click()
   '
   Command10.Enabled = False
   '
   TMENSA$ = Text2
   If TRIM$(TMENSA$) = "" Then Call MENSERR(24, "Falta Texto de Mensaje."): GoTo 99
   '
   DESTI1% = XVALO(Left$(Combo1.TEXT, 3))
   If DESTI1% < 1 Then Call MENSERR(24, "Falta Destinatario Principal del Mensaje."): GoTo 99
   '
   DESTI2% = XVALO(Left$(Combo3.TEXT, 3))
   DESTI3% = XVALO(Left$(Combo3.TEXT, 3))
   SOLICO% = Check1.Value
   IMEDIA% = Check2.Value
   '
   FHORPROG$ = AHORA$
   If Label4 <> "" Then FHORPROG$ = FETEXCOR1$(FECHANUM(Label4)) + " " + Text1.TEXT
   '
   LISTADEADJUN$ = ""
   For KKI% = 1 To List3.ListCount
      If LISTADEADJUN$ <> "" Then
         LISTADEADJUN$ = LISTADEADJUN$ + ";"
      End If
      LISTADEADJUN$ = LISTADEADJUN$ + List3.List(KKI% - 1)
   Next KKI%
   '
   DESTIX$ = USERNAM$(DESTI1%)
   If DESTI2% > 0 Then DESTIX$ = DESTIX$ + ", " & USERNAM$(DESTI2%)
   If DESTI3% > 0 Then DESTIX$ = DESTIX$ + ", " & USERNAM$(DESTI3%)
   '
   Call ENVIAMENSAJE(TMENSA$, ASUNTO$, DESTI1%, DESTI2%, DESTI3%, SOLICO%, IMEDIA%, FHORPROG$, LISTADEADJUN$)
   '
   Call COMUNI("Mensaje Enviado a '" + DESTIX$ + "'.")
Unload Me                       '  ' PARA IMPIAR LA PANTALLA
   '
99 Command10.Enabled = True
End Sub

Sub ENVIAMENSAJE(TMENSA$, ASUNTO$, DESTI1%, DESTI2%, DESTI3%, SOLICO%, IMEDIA%, FHORPROG$, Optional LISTADEADJUN$)
   '
   Call CREACAMPO("", "MENSAINT", "NUME_MENSA", 4, 4, 1)
   Call CREACAMPO("", "MENSAINT", "ASUN_MENSA", 10, 256)
   Call CREACAMPO("", "MENSAINT", "PARA_MENSA", 3, 2, 1)
   Call CREACAMPO("", "MENSAINT", "CCOP_MENSA", 3, 2, 1)
   Call CREACAMPO("", "MENSAINT", "TEXT_MENSA", 12, 0)
   Call CREACAMPO("", "MENSAINT", "SOLI_CONF", 3, 2)
   Call CREACAMPO("", "MENSAINT", "ENVIA_YA", 3, 2)            ' MARCA DE ENVIO INMEDIATO
   Call CREACAMPO("", "MENSAINT", "CONF_RECEP", 8, 0, 1)       ' FECHA Y HORA CONFIRMACION DE RECEPCION
   Call CREACAMPO("", "MENSAINT", "CONF_RECOP", 8, 0, 1)       ' FECHA Y HORA CONFIRMACION DE RECEPCION COPIA
   Call CREACAMPO("", "MENSAINT", "FHOR_PREPA", 8, 0, 1)       ' FECHA Y HORA QUE FUE PREPARADO
   Call CREACAMPO("", "MENSAINT", "USUA_PREPA", 3, 2, 1)       ' USUARIO QUE LO PREPARÓ
   Call CREACAMPO("", "MENSAINT", "FHOR_SOLENVI", 8, 0, 1)     ' FECHA Y HORA PROXIMO ENVIO PROGRAMADO
   Call CREACAMPO("", "MENSAINT", "FHOR_ULTENVI", 8, 0, 0)     ' FECHA Y HORA ULTIMO PRESENTACION
   Call CREACAMPO("", "MENSAINT", "FHOR_ULTENCO", 8, 0, 0)     ' FECHA Y HORA ULTIMO PRESENTACION
   Call CREACAMPO("", "MENSAINT", "LISADJUN", 12, 0)           ' LISTA DE ADJUNTOS
   '
   NUPROMEN& = 1 + XVALO(VALOBADA("MENSAINT", "MAX(NUME_MENSA)", "NUME_MENSA > 0", "NOMESS"))
   CONDI$ = "NUME_MENSA = " & NUPROMEN&
   Call CreaRegistro("MENSAINT", "NUME_MENSA", NUPROMEN&, "NOMESS")
   ' Label10 = TRIM$(Str$(NUPROMAI&))
   '
   Call ACTUREGISTRO("MENSAINT", "NUME_MENSA", CONDI$, NUPROMEN&, RAFE&, "NOMESS")
   Call ACTUREGISTRO("MENSAINT", "ASUN_MENSA", CONDI$, Left$(ASUNTO$, 256), RAFE&, "NOMESS")
   Call ACTUREGISTRO("MENSAINT", "PARA_MENSA", CONDI$, DESTI1%, RAFE&, "NOMESS")
   Call ACTUREGISTRO("MENSAINT", "CCOP_MENSA", CONDI$, DESTI2%, RAFE&, "NOMESS")
   Call ACTUREGISTRO("MENSAINT", "TEXT_MENSA", CONDI$, TMENSA$, RAFE&, "NOMESS")
   Call ACTUREGISTRO("MENSAINT", "ENVIA_YA", CONDI$, IMEDIA%, RAFE&, "NOMESS")
   Call ACTUREGISTRO("MENSAINT", "SOLI_CONF", CONDI$, SOLICO%, RAFE&, "NOMESS")
   Call ACTUREGISTRO("MENSAINT", "CONF_RECEP", CONDI$, FETEXCOR1$(FECHANUM("01/01/80")), RAFE&, "NOMESS")
   Call ACTUREGISTRO("MENSAINT", "FHOR_PREPA", CONDI$, AHORA$, RAFE&, "NOMESS")           ' F-HOR PREPARADO
   Call ACTUREGISTRO("MENSAINT", "USUA_PREPA", CONDI$, USNBR% Mod 100, RAFE&, "NOMESS")   ' USUARIO PREPARÓ
   Call ACTUREGISTRO("MENSAINT", "FHOR_SOLENVI", CONDI$, FHORPROG$, RAFE&, "NOMESS")       ' PROXIMO ENVIO
   Call ACTUREGISTRO("MENSAINT", "FHOR_ULTENVI", CONDI$, NUNCA$, RAFE&, "NOMESS")
   Call ACTUREGISTRO("MENSAINT", "FHOR_ULTENCO", CONDI$, NUNCA$, RAFE&, "NOMESS")
   Call ACTUREGISTRO("MENSAINT", "LISADJUN", CONDI$, LISTADEADJUN$, RAFE&, "NOMESS")
   '
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   IXS& = List3.ListIndex
   If IXS& >= 0 Then
     List3.RemoveItem ISX&
   End If
   Command2.Enabled = True
End Sub

Private Sub Command7_Click()
   Call CIERRARCH("*.*")
   Hide
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   DOCADJUN09.Show 1
   For U& = 1 To DOCADJUN09.List1.ListCount
      FIFUNA$ = TRIM$(DOCADJUN09.List1.List(U& - 1))
      If EXISTE%(TRIM$(FIFUNA$)) > 0 Then
         For JJ% = 1 To List3.ListCount
            If UCase$(List3.List(JJ% - 1)) = UCase$(FIFUNA$) Then GoTo 19
         Next JJ%
         List3.AddItem FIFUNA$
      End If
19 Next U&
   DOCADJUN09.List1.Clear
   Unload DOCADJUN09
   Command8.Enabled = True
End Sub

Private Sub Command9_Click()
   '
   Call SELECHO("SELFECHA")
   VDEV$ = VALACT1$("SELFECHA")
   If VDEV$ <> "" Then
       Label4.Caption = VDEV$
   End If

End Sub

Private Sub Form_Activate()
   If ATENTI% > 0 Then
     Left = 0
     Top = 0
     Call CENTRAFORM(Me)
     ATENTI% = 0
     Text2.SetFocus
   End If
End Sub

Private Sub Form_Load()
   '
    Call CENTRAFORM(Me)
    Call CARLIUS(Combo1)
    Call CARLIUS(Combo2)
    Call CARLIUS(Combo3)
    '
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

Private Sub List3_DblClick()
   DOCUACT$ = TRIM$(Mid$(List3.TEXT, 129))
   Call MUESTRADOC4(DOCUACT$)
End Sub

Sub MUESTRADOC4(DOCUACT$)
    '
    Call ShellExecute(Me.hwnd, "Open", DOCUACT$, "", "", 1)
    '
End Sub

Private Sub Text2_Change()
    UTIMETEC = Timer
End Sub
