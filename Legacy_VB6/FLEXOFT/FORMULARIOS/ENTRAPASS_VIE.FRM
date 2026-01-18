VERSION 5.00
Begin VB.Form ENTRAPASS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingresa Clave"
   ClientHeight    =   1110
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1110
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Height          =   720
      Left            =   2880
      Picture         =   "ENTRAPASS_VIE.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   550
   End
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
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   645
      Width           =   1575
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
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Password:"
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
      Left            =   -120
      TabIndex        =   4
      Top             =   690
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Username:"
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
      Left            =   -120
      TabIndex        =   3
      Top             =   285
      Width           =   1215
   End
End
Attribute VB_Name = "ENTRAPASS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ESTADO
'
Private Sub Command3_Click()
    Static VUSER$
    If VUSER$ = "" Then
        ' CARGA TABLA DE USUARIOS
        VUSER$ = LOADFILE$("USER.DAT")
        If VUSER$ = "" Then
            ' SI NO HAY FIRMANTES DEFINIDOS CANCELA EJECUCION
            CARPACT$ = REPLACAR(CurDir, "\", "/")
            Call MENSERR(24, "Imposible Registrar Firma Electronica.\   \Falta Tabla 'USER.DAT' en Carpeta\'" + CARPACT$ + "'")
            Call FINAL("")
        End If
    End If
    '
    CAUS% = Len(VUSER$) / 64
    For JJ% = 1 To CAUS%
       F0$ = Mid$(VUSER$, 64 * JJ% - 63, 64)
       IDENT$ = Mid$(F0$, 31, 12)
         If Left$(IDENT$, 1) = "@" Then IDENT$ = DESENCRI$(Mid$(F0$, 32, 11)) + " "
         If LTrim$(RTrim$(IDENT$)) = "" Then GoTo 19
       DENO$ = Left$(F0$, 30)
         If Left$(DENO$, 1) = "@" Then DENO$ = DESENCRI$(Mid$(F0$, 2, 29)) + " "
       CLAV$ = Mid$(F0$, 43, 12)
         If Left$(CLAV$, 1) = "@" Then CLAV$ = DESENCRI$(Mid$(F0$, 44, 11)) + " "
         CLAV$ = REPLACAR$(CLAV$, Chr$(0), Chr$(32))
       If UCase$(Text1) = TRIM$(UCase$(IDENT$)) Then
         If UCase$(Text2) = TRIM$(UCase$(CLAV$)) Then
           GoTo 99
         End If
       End If
19  Next JJ%
    '
40  Call MENSERR(24, "Usuario / Clave Incorrecta")
    Text1 = "": Text2 = ""
    '
99  Hide
    '
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    Text1.SetFocus
    On Error GoTo 0
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
         Text2.SetFocus
       ElseIf KeyAscii = 27 Then
         Unload Me
    End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
         Command3.SetFocus
       ElseIf KeyAscii = 27 Then
         Unload Me
    End If
End Sub
