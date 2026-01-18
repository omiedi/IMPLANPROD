VERSION 5.00
Begin VB.Form IngClave1 
   Caption         =   "Validacion de Clave"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3345
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   3345
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Copiar al Portapapeles"
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   4800
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   360
      TabIndex        =   8
      Top             =   4320
      Width           =   2775
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      MaxLength       =   6
      TabIndex        =   6
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      MaxLength       =   6
      TabIndex        =   5
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   3720
      Width           =   3015
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   5280
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Respuesta2"
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
      Left            =   1680
      TabIndex        =   14
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Respuesta1"
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
      TabIndex        =   13
      Top             =   2760
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderWidth     =   5
      X1              =   0
      X2              =   3480
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "1.-Envie por Mail a omiedi@yahoo.com.ar  la clave1 y la clave2 en ese orden."
      Height          =   735
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   3255
   End
   Begin VB.Label Label4 
      Caption         =   "Clave2"
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
      Left            =   1680
      TabIndex        =   11
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Clave1"
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
      TabIndex        =   10
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Si ya Obtuvo la contraseña ingresela en los casilleros correspondientes"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bienvenido: para conseguir la licencia de uso seguir los siguientes pasos:      "
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   3135
   End
End
Attribute VB_Name = "IngClave1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mcIni     As clsIni               ' guardar los datos en un archivo .ini

Dim INI_PATH      As String

Private Sub Command1_Click()

     TT1% = Len(Text1) + Len(Text2)
     Text3.SelStart = 0
     Text3.SelLength = TT1%
     Clipboard.SetText Text3.SelText

End Sub

Private Sub Command2_Click()
    Set mcIni = New clsIni
    INI_PATH = App.Path & "\config.ini"
    '
    If TRIM(Text1) = "" Or TRIM(Text2) = "" Or TRIM(Text4) = "" Or TRIM(Text5) = "" Then
       MsgBox "Revise la Información, Faltan Datos"
       Exit Sub
    End If
    '
    sas = CONTRACLA$(TRIM(Text3))
    If CONTRACLA$(TRIM(Text3)) = TRIM(Text6) Then
    
        With mcIni
            Call .writeValue(INI_PATH, "datos", "CLAVEIDA", .Encriptar(App.EXEName, Text3.TEXT, 1))
            Call .writeValue(INI_PATH, "datos", "CLAVEVUELTA", .Encriptar(App.EXEName, Text6.TEXT, 1))
        End With
        MsgBox "Ingreso Aceptado, Vuelva a Iniciar el Programa"
        End
        Close
        '
    Else
       MsgBox "Clave Inválida"
    End If
End Sub

Private Sub Form_Load()
'Para almacenar la unidad elegida del DriveListBox
Dim unidad As String
'Variable que retorna el Numero de serie del volumen
Dim Nserie As Long
'Para almacenar el sistema de archivos
Dim sistemaArchivos As String
'Para retornar el nombre del volumen
Dim volumen As String
'Para saber si funcionó o no la llamada a la función Api
Dim retorno As Long
'
volumen = String$(255, Chr$(0))

sistemaArchivos = String$(255, Chr$(0))

unidad = "C:\"

unidad = Left(unidad, 2) & "\"
'Llamamos a la función, pasandole las variables
retorno = GetVolumeInformation(unidad, volumen, Len(volumen), Nserie, 0, 0, sistemaArchivos, Len(sistemaArchivos))

'Mostramos los valores retornados y los mostramos en los controles Label
SERIE$ = "000000000000" & TRIM$(Str(Nserie))
SERIE1$ = Right(SERIE$, 6) 'LAS ULTIMAS 6 POSICIONES
SERIE2$ = Right(Mid$(SERIE$, 1, Len(SERIE$) - 6), 6) 'LAS ULTIMAS 6 POSICIONES DEL RESTO
Text1 = SERIE2$
Text2 = SERIE1$

Text3 = SERIE2$ & SERIE1$ ' text  no visible
'

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  End
  Close

End Sub

Private Sub Text4_Change()
   Text6 = TRIM$(Text4) & TRIM$(Text5)
   If Len(TRIM(Text4)) >= 6 Then
      On Error Resume Next
      Text5.SetFocus
      On Error GoTo 0
   End If
End Sub

Private Sub Text5_Change()
   Text6 = TRIM$(Text4) & TRIM$(Text5)
End Sub
