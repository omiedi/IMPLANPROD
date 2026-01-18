VERSION 5.00
Begin VB.Form FLWEBSE1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FLEXOFT Web Services"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Left            =   4200
      TabIndex        =   2
      Top             =   6240
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Registrarse"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Left            =   1680
      TabIndex        =   1
      Top             =   6240
      Width           =   2175
   End
   Begin VB.Label Label11 
      Caption         =   "Label11"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   2040
      TabIndex        =   12
      Top             =   4620
      Width           =   5265
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   7665
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Label Label10 
      Caption         =   $"FLWEBSE1.frx":0000
      Height          =   225
      Left            =   210
      TabIndex        =   11
      Top             =   7035
      Width           =   7260
   End
   Begin VB.Label Label9 
      Caption         =   $"FLWEBSE1.frx":0132
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   0
      TabIndex        =   10
      Top             =   5520
      Width           =   7470
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "- Recibir amplia información a cerca de nuevas aplicaciones disponibles, consul- tas y listadores."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   315
      TabIndex        =   9
      Top             =   3360
      Width           =   7050
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   $"FLWEBSE1.frx":0230
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   315
      TabIndex        =   8
      Top             =   2835
      Width           =   7050
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   " - Recibir 'Tips' o consejos de interés general basados en consultas fomuladas por otros usuarios de la comunidad FLEXOFT."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   315
      TabIndex        =   7
      Top             =   2310
      Width           =   7050
   End
   Begin VB.Label Label5 
      Caption         =   $"FLWEBSE1.frx":02CB
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   105
      TabIndex        =   6
      Top             =   3990
      Width           =   7575
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   " - Imprimir las respuestas recibidas, guardarlas y reenviarlas a otros usuarios den- tro y fuera de su LAN."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   315
      TabIndex        =   5
      Top             =   1785
      Width           =   7050
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "- Recibir respuestas a sus consultas telefónicas directamente en el escritorio de cada usuario."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1410
      Left            =   315
      TabIndex        =   4
      Top             =   1260
      Width           =   7050
   End
   Begin VB.Label Label2 
      Caption         =   $"FLWEBSE1.frx":0354
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   105
      TabIndex        =   3
      Top             =   525
      Width           =   7575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Bienvenidos a los FLEXOFT Web Services !!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   7575
   End
End
Attribute VB_Name = "FLWEBSE1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   '
   Call ENVIAMAIL("FLEXOFT", "asistec@flexoft.com.ar", "FLEXOFT Web Services", Label11.Caption)
   '
   UL$ = Left$(App.Path, 2)
   IDEPLI$ = SERIVOL$(UL$)
   Call SAVEFILE(App.Path + "\BKSEVOL.DAT", ENCRIPTA$(IDEPLI$))
   '
   ACUSER$ = LOADFILE$(App.Path + "\USER.DAT")
   Call SAVEFILE(App.Path + "\BKUSER.DAT", ACUSER$)
   '
   AHORA$ = Format$(Now, "dd/mm/yy")
   Call SAVEFILE(App.Path + "\SOLREGIS.DAT", ENCRIPTA$(AHORA$))
   '
   TTXX$ = "Su solicitud de registro y actualizacion de datos para los FLEXOFT WEB Services ha sido pasada a su programa de correos para su envío a asistec@flexoft.com.ar."
   TTXX$ = TTXX$ + vbCrLf + vbCrLf + "En breve recibirá confirmación y aceptación de su solicitud."
   TTXX$ = TTXX$ + vbCrLf + vbCrLf + "Caso contrario favor comunicarse telefónicamente al (54-11) 4524-1284 / 4522-2048 interno 32 / 33 en horario de 9:00 a 18:00 hs (GMT + 3)."
   TTXX$ = TTXX$ + vbCrLf + vbCrLf + "Muchas gracias por ayudarnos a mejorar nuestro servicio de soporte."
   TTXX$ = TTXX$ + vbCrLf + vbCrLf + "Por favor, re-inicie ahora la aplicación FLEXOFT."
   MsgBox TTXX$, vbExclamation, "FLEXOFT WEB Services"
   Close: End
   '
   Unload Me
End Sub

Private Sub Command2_Click()
   Unload Me
End Sub
