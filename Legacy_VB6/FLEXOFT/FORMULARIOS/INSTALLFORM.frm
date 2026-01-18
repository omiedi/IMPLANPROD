VERSION 5.00
Begin VB.Form INSTALLFORM 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Instalación FLEXOFT"
   ClientHeight    =   13500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   18000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   13500
   ScaleWidth      =   18000
   ShowInTaskbar   =   0   'False
   Begin VB.FileListBox File1 
      Height          =   3015
      Left            =   7800
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Para Cancelar la Instalación FLEXOFT,  Haga  un  Doble-Click  AQUI  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   975
      Left            =   7560
      TabIndex        =   4
      Top             =   7200
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Para Configurar la Estación de Trabajo, Instalando  Drivers  y  Controladores Locales,  Haga   Doble-Click  AQUI  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   1455
      Left            =   7560
      TabIndex        =   3
      Top             =   5160
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Para una Instalación y Configuración Completa Haga Doble-Click AQUI  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   975
      Left            =   7560
      TabIndex        =   2
      Top             =   3480
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Instalación de FLEXOFT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   7935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Instalación de FLEXOFT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   7935
   End
   Begin VB.Image Image1 
      Height          =   13500
      Left            =   0
      Picture         =   "INSTALLFORM.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   18000
   End
End
Attribute VB_Name = "INSTALLFORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ULORI$

Private Sub Form_Load()
   '
   Width = Screen.Width
   Height = Screen.Height
   Image1.Width = Width
   Image1.Height = Height
   '
   Label1.Top = Label2.Top + 60
   Label1.Left = Label2.Left + 80
   '
   Label3.Left = Width - Label3.Width - 400
   Label4.Left = Label3.Left
   Label5.Left = Label3.Left
   '
   Show
   Refresh
   If Screen.Width < 11800 Then
     CRLF$ = Chr$(13) + Chr$(10)
     TTXX$ = "Para Ejecutar los Sistemas FLEXOFT, el Area de" + CRLF$
     TTXX$ = TTXX$ + "Escritorio debe configurarse a 800 x 600 pix." + CRLF$
     TTXX$ = TTXX$ + "o mayor." + CRLF$ + CRLF$
     TTXX$ = TTXX$ + "Reconfigure el Area de Escritorio, y vuelva a" + CRLF$
     TTXX$ = TTXX$ + "iniciar la instalación de FLEXOFT."
     MsgBox TTXX$, 48, "Area de Escritorio"
     Close
     End
     Exit Sub
   End If
   '
   ADVIERTE.Label3 = ""
   ADVIERTE.Show 1
   If Val(ADVIERTE.Label3) <> 1 Then
     Close
     End
     Exit Sub
   End If
   Label3.Visible = True
   Label4.Visible = True
   Label5.Visible = True
   
End Sub

Private Sub Label3_dblclick()
   '
10 On Error GoTo 0
   SELUDESTI.Show 1
   If Trim$(SELUDESTI.Label3) = "" Then
     Close
     End
     Exit Sub
   End If
   '
   UDESTI$ = Trim$(Left$(SELUDESTI.Label3, 1))
   Call ULOGINSTA
   If ULORI$ <> "" Then
     CUDI$ = CurDir
     '
20   CARPEINS$ = UCase$(Trim$(InputBox$("Carpeta de Instalación:", "Configurar Instalación", "FLEXOFT")))
     If CARPEINS$ = "" Then Close: End
     If Len(CARPEINS$) > 8 Then
        MsgBox "Carpeta no Válida", 48
        GoTo 20
     End If
     For kki% = 1 To Len(CARPEINS$)
       KKJ% = Asc(Mid$(CARPEINS$, kki%))
       If KKJ% < 48 Or (KKJ% > 57 And KKJ% < 65) Or KKJ% > 90 Then
         MsgBox "Caracteres no Válidos", 48
         GoTo 20
       End If
     Next kki%
     '
     On Error Resume Next
     MkDir UDESTI$ + ":\" + CARPEINS$
     MkDir UDESTI$ + ":\" + CARPEINS$ + "\WINDOWS"
     MkDir UDESTI$ + ":\" + CARPEINS$ + "\DATOS"
     On Error GoTo 0
     '
     On Error Resume Next
     ChDir UDESTI$ + ":\" + CARPEINS$
     If Err Then
       MsgBox "Error Crítico de Acceso a la Unidad " + UDESTI$ + ":\" + Chr$(13) + Chr$(10) + "Imposible Crear Carpetas de Trabajo FLEXOFT.", 16, "Instalación de FLEXOFT"
       GoTo 10
     End If
     '
     On Error GoTo 0
     ChDir ULORI$ + ":\FX_SETUP"
     Shell ULORI$ + ":\FX_SETUP\SETUP.EXE", vbNormalFocus
     '
'MsgBox "1" & ULORI$ + ":\FX_SYSBAS"
     File1.Path = ULORI$ + ":\FX_SYSBAS"
'MsgBox 2 & File1.Path
     File1.Pattern = "*.*"
'MsgBox 3 & "   " & File1.ListCount
     For U& = 1 To File1.ListCount
       ANAME$ = File1.List(U& - 1)
       FileCopy ULORI$ + ":\FX_SYSBAS\" + ANAME$, UDESTI$ + ":\" + CARPEINS$ + "\WINDOWS\" + ANAME$
       On Error Resume Next
       SetAttr UDESTI$ + ":\" + CARPEINS$ + "\WINDOWS\" + ANAME$, vbNormal
       On Error GoTo 0
     Next U&
     '
'MsgBox "7" & ULORI$ + ":\FX_PROGRAMS"
     File1.Path = ULORI$ + ":\FX_PROGRAMS"
'MsgBox 8 & File1.Path
     File1.Pattern = "*.*"
'MsgBox 9 & "   " & File1.ListCount
     For U& = 1 To File1.ListCount
       ANAME$ = File1.List(U& - 1)
       On Error Resume Next
       FileCopy ULORI$ + ":\FX_PROGRAMS\" + ANAME$, UDESTI$ + ":\" + CARPEINS$ + "\WINDOWS\" + ANAME$
       If Err Then
         MsgBox "Imposible Instalar '" + ULORI$ + ":\FX_PROGRAMS\" + ANAME$ + "'"
       End If
       On Error GoTo 0
       On Error Resume Next
       SetAttr UDESTI$ + ":\" + CARPEINS$ + "\WINDOWS\" + ANAME$, vbNormal
       On Error GoTo 0
     Next U&
     '
80   MsgBox "Felicitaciones !!!" + Chr$(13) + Chr$(10) + Chr$(13) + Chr$(10) + "Ha Completado la Instalación FLEXOFT con éxito."
     '
     CRLF$ = Chr$(13) + Chr$(10)
     TTXX$ = ""
     TTXX$ = TTXX$ + "Si ha instalado FLEXOFT en red LAN, recuerde" + CRLF$
     TTXX$ = TTXX$ + "que la conexión a  '" + UDESTI$ + ":'  debe marcarse como" + CRLF$
     TTXX$ = TTXX$ + "'conectar de nuevo al iniciar la sesión'." + CRLF$
     TTXX$ = TTXX$ + CRLF$
     TTXX$ = TTXX$ + "No fue posible la instalación automatica en su" + CRLF$
     TTXX$ = TTXX$ + "escritorio. Para facilitar la ejecución se reco-" + CRLF$
     TTXX$ = TTXX$ + "mienta que cree un 'Nuevo Acceso Directo' a" + CRLF$
     TTXX$ = TTXX$ + UDESTI$ + ":\" + CARPEINS$ + "\WINDOWS\modulo_inicial.EXE" + CRLF$
     '
     MsgBox TTXX$, 48, "Instalación Completada"
     Close: End
     '
   End If
End Sub

Private Sub Label4_DblClick()
   Call ULOGINSTA
   If ULORI$ <> "" Then
     CUDI$ = CurDir
     ChDir ULORI$ + ":\FX_SETUP"
     Shell ULORI$ + ":\FX_SETUP\SETUP.EXE", vbNormalFocus
   End If
End Sub

Private Sub Label5_DblClick()
   Close
   End
End Sub

Sub ULOGINSTA()
   '
   ULORI$ = ""
   aa1$ = CurDir
   On Error Resume Next
   For KI% = 68 To 90    ' EMPIEZA DESDE D:
     TTXX$ = Chr$(KI%) + ":\FX_SYSBAS\FLEXPRFS.FLX"
     LOAR = FileLen(TTXX$)
     If LOAR > 0 Then
       ULORI$ = Chr$(KI%)
       GoTo 90
     End If
   Next KI%
   On Error GoTo 0
10 CRLF$ = Chr$(13) + Chr$(10)
   TTXX$ = "Imposible determinar la Unidad de Instalacion." + CRLF$ + CRLF$
   TTXX$ = TTXX$ + "Asegurese de que la unidad CD-ROM en que se" + CRLF$
   TTXX$ = TTXX$ + "encuentra el CD de Instalación, esté conecta-" + CRLF$
   TTXX$ = TTXX$ + "da como Unidad Local o de Red, mediante letra" + CRLF$
   TTXX$ = TTXX$ + "igual o superior a 'D:'."
   MsgBox TTXX$, 16, "Instalación de FLEXOFT"
   
   '
90 On Error GoTo 0
   On Error Resume Next
   ChDir aa1$
   On Error GoTo 0
   '
End Sub
