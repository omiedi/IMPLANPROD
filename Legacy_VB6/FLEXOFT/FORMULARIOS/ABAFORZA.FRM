VERSION 5.00
Begin VB.Form ABAFORZA 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FLEXOFT - Bajada Forzada de la WEB"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   9720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   400
      Left            =   0
      ScaleHeight     =   405
      ScaleWidth      =   435
      TabIndex        =   4
      Top             =   0
      Width           =   430
      Begin VB.Image Image3 
         Height          =   240
         Left            =   105
         Picture         =   "ABAFORZA.frx":0000
         ToolTipText     =   "Doble-Click para Des-Habilitar Web Services en esta Estación de Trabajo"
         Top             =   105
         Width           =   240
      End
   End
   Begin VB.ListBox List3 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2460
      IntegralHeight  =   0   'False
      ItemData        =   "ABAFORZA.frx":0E42
      Left            =   105
      List            =   "ABAFORZA.frx":0E44
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   1215
      Width           =   9480
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nombre del Archivo              Fecha    Hora   Destino-1   Destino-2   Ctrl"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1005
      Width           =   9495
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
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
      Top             =   525
      Width           =   9500
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   10080
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   105
      TabIndex        =   1
      Top             =   315
      Width           =   9500
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
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
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   9500
   End
End
Attribute VB_Name = "ABAFORZA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Form_Load()
   '
   On Error Resume Next
   Caption = "WEB Services FLEXOFT - " + Format(FileDateTime(App.Path + "\FLX_BRWS.EXE"), "dd/mm/yy  hh:mm")
   On Error GoTo 0
   '
   Call LEEDATSYS
   Call CREABACONLEC
   '
   Label1 = NEPLI$ + " (" + IDEPLI$ + ")"
      TTXX$ = Trim$(Str$(USERNU%)) + " - " + Trim$(USERNA$) + " (" + Trim$(USERID$) + ")"
   Label4 = TTXX$
   Label8 = CARPEDOWN$
   '
End Sub

Private Sub List3_DblClick()
    WEBROWSR.Timer2.Enabled = False
    ARNAME$ = UCase$(Trim$(Left$(List3.Text, 32)))
    UDEST1$ = UCase$(Trim$(Mid$(List3.Text, 49, 12)))
    UDEST2$ = UCase$(Trim$(Mid$(List3.Text, 61, 12)))
    If (USERNU% Mod 100) <> 1 Then            ' NUMERO DE USUARIO
      If UCase$(USERID$) <> UDEST1$ Then      ' IDENTIFICACION DE USUARIO
        If UCase$(USERID$) <> UDEST2$ Then    ' IDENTIFICACION DE USUARIO
          MsgBox "Imposible Acceder a la Petición. Solo Alguno de los Destinatarios o el Supervisor FLEXOFT Pueden Abrir / Bajar el Archivo Elegido."
          GoTo 99
        End If
      End If
    End If
    JJ1& = BAJADELAWEB(ARNAME$)
    If JJ1& <> 0 Then
       MsgBox "Error al Bajar Archivo de la WEB"
       Exit Sub
    End If
    '
    ARFULLNAME$ = AJUSTI$(List3.Text, 48)
    If InStr(ARNAME$, ".PDF") > 0 Then
        Call WEBROWSR.MUESTRAPDF(ARNAME$, ARFULLNAME$)
      ElseIf InStr(ARNAME$, ".DOC") > 0 Then
        Call WEBROWSR.MUESTRADOC(ARNAME$, ARFULLNAME$)
      ElseIf InStr(ARNAME$, ".XLS") > 0 Then
        Call WEBROWSR.MUESTRADOC(ARNAME$, ARFULLNAME$)
      Else
        Call WEBROWSR.INSTALACON(ARFULLNAME$)
    End If
99  WEBROWSR.Timer2.Enabled = True
End Sub
'
