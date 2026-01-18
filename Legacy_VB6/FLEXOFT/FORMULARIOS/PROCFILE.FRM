VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form PROCFILE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Procesamiento de Mensajes FLEXOFT"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11715
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   11715
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
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
      Height          =   1290
      Left            =   5040
      TabIndex        =   8
      Top             =   120
      Width           =   6525
      Begin VB.CommandButton Command11 
         Caption         =   "Pre Instalar Todos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   850
         Left            =   5400
         TabIndex        =   13
         Top             =   300
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   840
         Width           =   2655
      End
      Begin VB.CommandButton Command6 
         Caption         =   "ok"
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
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   840
         Width           =   375
      End
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
         Height          =   420
         Left            =   120
         TabIndex        =   9
         Top             =   300
         Width           =   5140
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Re-Enviar a:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -600
         TabIndex        =   12
         Top             =   900
         Width           =   2085
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   4725
      Top             =   0
   End
   Begin VB.Frame Frame1 
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
      Height          =   930
      Left            =   105
      TabIndex        =   1
      Top             =   480
      Width           =   4830
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
         Height          =   500
         Left            =   3450
         Picture         =   "PROCFILE.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Salir"
         Top             =   280
         Width           =   600
      End
      Begin VB.CommandButton Command7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   720
         Picture         =   "PROCFILE.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Enviar Confirmacion de  Lectura"
         Top             =   280
         Width           =   600
      End
      Begin VB.CommandButton Command8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   1350
         Picture         =   "PROCFILE.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Imprimir el Documento"
         Top             =   280
         Width           =   600
      End
      Begin VB.CommandButton Command9 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   1980
         Picture         =   "PROCFILE.frx":0ABE
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Guardar como Archivo"
         Top             =   280
         Width           =   600
      End
      Begin VB.CommandButton Command10 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   2610
         Picture         =   "PROCFILE.frx":0DC8
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Enviar como Adjunto por Mail"
         Top             =   280
         Width           =   600
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label5 
      Height          =   255
      Left            =   600
      TabIndex        =   15
      Top             =   6240
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.Label Label4 
      Height          =   375
      Left            =   600
      TabIndex        =   14
      Top             =   5640
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.Label Label1 
      Height          =   330
      Left            =   0
      TabIndex        =   7
      Top             =   5040
      Visible         =   0   'False
      Width           =   6315
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4800
   End
End
Attribute VB_Name = "PROCFILE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image3_DblClick()
   If Left$(UCase$(LUCOM$), 2) = "C:" Then
      If WEBSE$ = "ENABLED" Then
          WEBSE$ = "DISABLED"
        Else
          WEBSE$ = "ENABLED"
      End If
      Call SAVEFILE(LUCOM$ + "WEB_SER.TXT", WEBSE$)
      MsgBox "Re-Configuración Completada - Re-Inicie la Aplicación."
      Close: End
   End If
End Sub

Private Sub Command10_Click()
   Command10.Enabled = False
   Call SHOWPDF.Command10_Click
   Command10.Enabled = False
End Sub

Private Sub Command5_Click()
   Unload Me
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   If Trim$(Combo1.Text) <> "" Then
      ORIID$ = Trim$(UCase$(USERID$))
      DESTID$ = Trim$(Mid$(Combo1.Text, 65, 12))
      If DESTID$ <> "" Then
        DESTID$ = AJUSTI$(DESTID$, 12)
        IDME$ = Label5
        '
        Label8 = CARPEDOWN$
        '
        LISNOVE1$ = ""
        LISNOVE$ = LOADFILE$(CARPEDOWN$ + "NOPENPRO.TXT")
        For u& = 1 To Len(LISNOVE$) Step 128
           ARCHAI$ = Mid$(LISNOVE$, u&, 128)
           If Trim$(Mid$(ARCHAI$, 103, 4)) = IDME$ Then
              ORIUS$ = Trim$(UCase$(Mid$(ARCHAI$, 75, 12)))
              CONCO$ = Trim$(UCase$(Mid$(ARCHAI$, 107, 12)))
              If ORIUS$ = ORIID$ Or CONCO$ = ORIID$ Or ORIUS$ = "" Then
                 Mid$(ARCHAI$, 75, 12) = DESTID$
                 Mid$(ARCHAI$, 107, 12) = Space$(12)   ' QUEDA EN BLANCO
                 LISNOVE1$ = LISNOVE1$ + ARCHAI$
              End If
           End If
        Next u&
        '
        While Len(LISNOVE$) Mod 128 <> 0: LISNOVE$ = LISNOVE$ + " ": Wend
        LISNOVE$ = LISNOVE$ + LISNOVE1$
        While Len(LISNOVE$) Mod 128 <> 0: LISNOVE$ = LISNOVE$ + " ": Wend
        Call SAVEFILE(CARPEDOWN$ + "NOPENPRO.TXT", LISNOVE$)
        '
      End If
   End If
   '
   Command6.Enabled = True
End Sub

Sub Command7_Click()
   Command7.Enabled = False
   Call WEBROWSR.CONFIRMALEC(Label1)
   Command7.Enabled = True
End Sub
'
Private Sub Command8_Click()
   Command8.Enabled = False
   Call SHOWPDF.Command8_Click
   Command8.Enabled = False
End Sub

Private Sub Command9_Click()
   Command9.Enabled = False
   Call SHOWPDF.Command9_Click
   Command9.Enabled = False
End Sub

Private Sub Form_Load()
   '
   Combo1.Clear
   LFI$ = LOADFILE$("C:\FLEXOFT\FLEXRAM.DRV")
   USERNU% = CVI(Mid$(LFI$, 9, 2)) Mod 100
   If USERNU% < 1 Then USERNU% = 1
   '
   VUSER$ = LOADFILE$("USER.DAT")
   For KKI% = 1 To Len(LFI$) / 64
     If KKI% <> USERNU% Then
       F0$ = Mid$(VUSER$, 64 * KKI% - 63, 64)
       USID$ = Mid$(F0$, 31, 12)
          If Left$(USID$, 1) = "@" Then USID$ = DESENCRI$(Mid$(F0$, 32, 11)) + " "
       USNA$ = Left$(F0$, 30)
          If Left$(USNA$, 1) = "@" Then USNA$ = DESENCRI$(Mid$(F0$, 2, 29)) + " "
       If Trim$(USID$) <> "" Then
         If Trim$(USNA$) <> "" Then
           Combo1.AddItem AJUSTI$(USNA$, 64) + USID$
         End If
       End If
     End If
   Next KKI%
   '
End Sub

Private Sub Timer1_Timer()
   If Visible = True Then
      Me.ZOrder 0
      AppActivate Caption
   End If
End Sub

Private Sub List1_DblClick()
    WEBROWSR.Timer2.Enabled = False
    ARNAME$ = UCase$(Trim$(Left$(List1.Text, 32)))
    ARFULLNAME$ = AJUSTI$(List1.Text, 48)
    If InStr(ARNAME$, ".EXE") > 0 Then
          Call WEBROWSR.INSTALACON(ARFULLNAME$)
          ' ELIMINA DE LA LISTA DE SELECCION
          ARFU$ = AJUSTI$(ARFULLNAME$, 48)
65        For KKIL& = 1 To WEBROWSR.List3.ListCount
            If Left$(List1.List(KKIL& - 1), 48) = ARFU$ Then
              List1.RemoveItem KKIL& - 1
              List1.Refresh
              GoTo 65
            End If
          Next KKIL&
       Else
          MsgBox "Imposible Instalar '" & ARNAME$ & "'."
    End If
    WEBROWSR.Timer2.Enabled = True
End Sub

