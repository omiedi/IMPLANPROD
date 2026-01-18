VERSION 5.00
Begin VB.Form FORBUSKE 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Busqueda de Archivos en Red LAN"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Opciones de Busqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      Left            =   6720
      TabIndex        =   3
      Top             =   320
      Width           =   4095
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
         Left            =   480
         TabIndex        =   8
         Top             =   1560
         Width           =   3375
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
         Height          =   285
         Left            =   2640
         TabIndex        =   6
         Text            =   "*.*"
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Por Contenido"
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
         Left            =   360
         TabIndex        =   5
         Top             =   600
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Por Nombre de Archivo"
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
         Left            =   360
         TabIndex        =   4
         Top             =   360
         Value           =   -1  'True
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Filtro:"
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
         Left            =   2400
         TabIndex        =   7
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Texto a Buscar:"
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
         Left            =   240
         TabIndex        =   9
         Top             =   1320
         Width           =   2415
      End
   End
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   7440
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   2430
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3840
      Left            =   120
      TabIndex        =   10
      ToolTipText     =   "Doble Click para Abrir Archivo"
      Top             =   2400
      Width           =   11655
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Doble Click para Detener Exploracion de Carpetas"
      Top             =   405
      Width           =   6300
   End
   Begin VB.CommandButton Command1 
      Height          =   1880
      Left            =   11040
      Picture         =   "FORBUSKE.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Inicia Busqueda"
      Top             =   390
      Width           =   690
   End
   Begin VB.Label Label5 
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
      Left            =   720
      TabIndex        =   13
      Top             =   2040
      Width           =   5725
   End
   Begin VB.Label Label4 
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
      Left            =   120
      TabIndex        =   12
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Carpetas a Explorar:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   135
      Width           =   3930
   End
End
Attribute VB_Name = "FORBUSKE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    
Dim FINEXPLORA%

Private Sub Command1_Click()
   '
   Command1.Enabled = False
   Screen.MousePointer = 11
   ABUSKE$ = TRIM$(UCase$(Text2))
   If ABUSKE$ = "" Then
      Call MENSERR(24, "Introduzca Texto a Buscar")
      GoTo 99
   End If
   '
   MODOBU% = 1: If Option2.Value = True Then MODOBU% = 2
   FINEXPLORA% = 0
   '
   If List1.ListCount < 1 Then
      List1.Clear: List2.Clear
      ILIST1& = 0
      '
      For LETRA% = 67 To 90
          ESTARUT$ = Chr$(LETRA%) + ":\"
          Label4 = ESTARUT$
          OPENFORMS = DoEvents
          List1.AddItem ESTARUT$
          '
          ' EXPLORA LOS DIRECTORIOS LOCALES
          Do While ILIST1& < List1.ListCount
             RUTAX$ = List1.List(ILIST1&)
             On Error Resume Next
             NOMBREX$ = Dir(RUTAX$, 16) ' Recupera la primera entrada.
             If Err Then
                On Error GoTo 0
                GoTo 6
             End If
             On Error GoTo 0
             Do While NOMBREX$ <> "" ' Inicia el bucle.
               ' Ignora el directorio actual y el que lo abarca.
               If NOMBREX$ <> "." And NOMBREX$ <> ".." Then
                  ' Utiliza comparación a nivel de bits para asegurarse de que MiNombre es un directorio.
                  On Error Resume Next
                  If (GetAttr(RUTAX$ & NOMBREX$) And vbDirectory) = vbDirectory Then
                     If Err Then
                        On Error GoTo 0
                        GoTo 45
                     End If
                     On Error GoTo 0
                     List1.AddItem RUTAX$ + NOMBREX$ + "\"
                     List1.ListIndex = List1.ListCount - 1
                  End If ' solamente si representa un directorio.
               End If
45             NOMBREX$ = Dir ' Obtiene siguiente entrada.
               Label5 = NOMBREX$
               OPENFORMS = DoEvents
               If FINEXPLORA% > 0 Then
                  MsgBox "Fin Exploración de Carpetas. Paso a Revision de Archivos."
                  GoTo 60
               End If
             Loop
6            ILIST1& = ILIST1& + 1
          Loop
      Next LETRA%
   End If
   '
60 ' CARGA EN LIST2 LOS ARCHIVOS LOCALES
   List2.Clear
   For i& = 1 To List1.ListCount
      List1.ListIndex = i& - 1
      On Error Resume Next
      File1.Path = List1.List(i& - 1)
      If Err Then
         On Error GoTo 0
         GoTo 19
      End If
      On Error GoTo 0
      File1.Pattern = Text1
      ESTARUTA$ = File1.Path
      If Right$(ESTARUTA$, 1) <> "\" Then ESTARUTA$ = ESTARUTA$ + "\"
      Label4 = Left$(File1.Path, 3): Label5 = File1.Path
      For J& = 1 To File1.ListCount
         FINAME$ = TRIM$(UCase$(File1.List(J& - 1)))
         Label5 = FINAME$: Label5.Refresh
         If MODOBU% = 1 Then
              TEBUS$ = FINAME$
            ElseIf MODOBU% = 2 Then
              TEBUS$ = UCase$(LOADFILE$(ESTARUTA$ + FINAME$))
         End If
         '
         If InStr(TEBUS$, ABUSKE$) > 0 Then
            TTXX$ = AJUSTI$(FINAME$, 64)
            TTXX$ = TTXX$ + ESTARUTA$ + FINAME$
            List2.AddItem TTXX$
            OPENFORMS = DoEvents
         End If
17    Next J&
19 Next i&
   '
99 Screen.MousePointer = 1
   Command1.Enabled = True
End Sub

Private Sub List1_DblClick()
   FINEXPLORA% = 1
End Sub

Private Sub List2_DblClick()
    DOCUACT$ = TRIM$(Mid$(List2.TEXT, 65))
    Call ShellExecute(Me.hwnd, "Open", DOCUACT$, "", "", 1)
End Sub
