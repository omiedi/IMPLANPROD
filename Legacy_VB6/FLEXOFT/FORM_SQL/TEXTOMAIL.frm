VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Begin VB.Form TEXTOMAIL 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Texto en Mail"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   7860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Obs Pred."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   260
      Left            =   960
      TabIndex        =   14
      ToolTipText     =   "Observaciones Predefiniidas"
      Top             =   2865
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2400
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   12
      Top             =   720
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00404040&
      Height          =   8520
      Left            =   6960
      ScaleHeight     =   8460
      ScaleWidth      =   1005
      TabIndex        =   8
      Top             =   0
      Width           =   1065
      Begin VB.CommandButton Command22 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   130
         Picture         =   "TEXTOMAIL.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Cancela Observaciones"
         Top             =   120
         Width           =   600
      End
      Begin VB.CommandButton Command14 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   130
         Picture         =   "TEXTOMAIL.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Aprueba, Observaciones"
         Top             =   2040
         Width           =   600
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "TEXTOMAIL.frx":0454
         Top             =   2640
         Width           =   1515
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Selecc Todo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   260
      Left            =   5700
      TabIndex        =   6
      Top             =   2865
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Pegar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   260
      Left            =   4485
      TabIndex        =   5
      Top             =   2865
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Copiar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   260
      Left            =   3240
      TabIndex        =   4
      Top             =   2865
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Cortar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   260
      Left            =   2055
      TabIndex        =   3
      Top             =   2865
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.TextBox Text28 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   105
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   720
      Width           =   6780
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00E0E0E0&
      Height          =   435
      Left            =   105
      ScaleHeight     =   375
      ScaleWidth      =   6720
      TabIndex        =   0
      Top             =   210
      Width           =   6780
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Texto en Mail"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   105
         TabIndex        =   1
         Top             =   105
         Width           =   5370
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "TEXTOMAIL.frx":2376
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   60
      OleObjectBlob   =   "TEXTOMAIL.frx":25AA
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   210
      TabIndex        =   11
      Top             =   4725
      Width           =   750
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   210
      TabIndex        =   7
      Top             =   5250
      Width           =   795
   End
End
Attribute VB_Name = "TEXTOMAIL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
      Command1.Enabled = False
      If EXISTE%(LUDAT$ + "TEXTOBS2" + ".RFS") = 0 Then
         Call COMUNI("No Desiponible (Solicite Actualización de Texto de Observaciones )")
         GoTo 99
      End If
      Call CREACAMPO("", "TAB_TEXTO", "CODINT_OBS", 3, 0, 1)
      Call CREACAMPO("", "TAB_TEXTO", "TITU_OBS", 10, 64)
      Call CREACAMPO("", "TAB_TEXTO", "TEXTO_OBS", 12, 0)
      Call TRALOCAL("TEXTOBS2", "")

      CONDI$ = "CODINT_OBS > 0 ORDER BY CODINT_OBS"
      Call CARGALISEL("LISTTITU", "TAB_TEXTO", "CODINT_OBS/F5;TITU_OBS/A48", CONDI$, "NOMESS/DISTINCT")
      Call SELECHO("LISTTITU")
      If VALACT1$("LISTTITU") = "" Then GoTo 99 ' COMALTER
      '
      CODITITU% = XVALO(VALACT1$("LISTTITU"))
      CONDI$ = "CODINT_OBS = " & CODITITU%
      TEXTOX$ = VALOBADA("TAB_TEXTO", "TEXTO_OBS", CONDI$, "NOMESS")
      If TRIM$(Text28) = "" Then Text28 = TEXTOX$: GoTo 99
      OPX% = COMALTER%("Textos Predefinidos\\Agregar al Final\Sobre-Escribir")
      If OPX% < 1 Or OPX% > 2 Then GoTo 99
      If OPX% = 1 Then
           Text28 = Text28 & vbCrLf & TEXTOX$
      ElseIf OPX% = 2 Then
           Text28 = TEXTOX$
      End If
      
99    Command1.Enabled = True

End Sub

Private Sub Command14_Click()
'   Label1 = "OK"
'   Hide
    Call GRACONTROL("PRESPAHP", "TEXTMAIL", Text28)
End Sub

Private Sub Command22_Click()
   Unload Me
End Sub

Private Sub Command17_Click()
   COMIE = Text28.SelStart
   LOSEL = Text28.SelLength
   Clipboard.SetText Mid$(Text28.TEXT, Text28.SelStart + 1, Text28.SelLength)
   Text28.TEXT = Left$(Text28.TEXT, COMIE) + Mid$(Text28.TEXT, COMIE + LOSEL + 1)
End Sub
'
Private Sub Command8_Click()
   Clipboard.SetText Mid$(Text28.TEXT, Text28.SelStart + 1, Text28.SelLength)
End Sub
'
Private Sub Command6_Click()
   COMIE = Text28.SelStart
   LOSEL = Text28.SelLength
   Text28.TEXT = Left$(Text28.TEXT, COMIE) + Clipboard.GetText + Mid$(Text28.TEXT, COMIE + LOSEL + 1)
End Sub
'
Private Sub Command5_Click()
   Text28.SelStart = 0
   Text28.SelLength = Len(Text28.TEXT)
   Text28.SetFocus
End Sub
'

Private Sub Form_Load()
   '
   Call APLICACIONSKINS(Me, Picture5)
   '
   Text28 = Control("PRESPAHP", "TEXTMAIL")
   
End Sub

