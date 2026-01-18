VERSION 5.00
Begin VB.Form WINDEMO2 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuracion de Pantalla"
   ClientHeight    =   6660
   ClientLeft      =   150
   ClientTop       =   405
   ClientWidth     =   7530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   7530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   645
      Left            =   630
      Picture         =   "WINDEMO2.frx":0000
      ScaleHeight     =   645
      ScaleWidth      =   540
      TabIndex        =   17
      Top             =   5145
      Width           =   540
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   645
      Left            =   210
      Picture         =   "WINDEMO2.frx":030A
      ScaleHeight     =   645
      ScaleWidth      =   540
      TabIndex        =   4
      Top             =   315
      Width           =   540
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Continuar Demostración"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   5040
      TabIndex        =   2
      Top             =   5985
      Width           =   2220
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Imprimir Instrucciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   2625
      TabIndex        =   1
      Top             =   5985
      Width           =   2220
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Configurar Monitor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   210
      TabIndex        =   0
      Top             =   5985
      Width           =   2220
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "10.- Para ayuda llámenos al: 54-11-4541-1515."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Index           =   9
      Left            =   1365
      TabIndex        =   16
      Top             =   5250
      Width           =   5790
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "9.- Reinicie la demostración."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Index           =   8
      Left            =   1470
      TabIndex        =   15
      Top             =   4935
      Width           =   5685
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "8.- Click sobre todos los botones Aceptar que  van                            apareciendo."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Index           =   4
      Left            =   1470
      TabIndex        =   14
      Top             =   4410
      Width           =   5685
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "7.- Ajustar Area de Escritorio a 800 por 600 píxeles."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Index           =   7
      Left            =   1470
      TabIndex        =   13
      Top             =   4095
      Width           =   5685
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "6.- Click sobre Configuración."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Index           =   6
      Left            =   1470
      TabIndex        =   12
      Top             =   3780
      Width           =   5685
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "5.- Doble Click sobre el ícono Pantalla."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Index           =   5
      Left            =   1470
      TabIndex        =   11
      Top             =   3465
      Width           =   5685
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "4.- Doble Click sobre el ícono Panel de Control."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Index           =   3
      Left            =   1470
      TabIndex        =   10
      Top             =   3150
      Width           =   5685
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "3.- Doble Click sobre el ícono Mi PC."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Index           =   2
      Left            =   1470
      TabIndex        =   9
      Top             =   2835
      Width           =   5685
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "2.- Sitúese sobre el Escritorio Principal."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Index           =   1
      Left            =   1470
      TabIndex        =   8
      Top             =   2520
      Width           =   5685
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "1.- Cancele la Demostración haciendo Click sobre el botón                 'Configurar Monitor'."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Index           =   0
      Left            =   1470
      TabIndex        =   7
      Top             =   1995
      Width           =   5685
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Siga las siguientes Instrucciones:"
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
      Left            =   1155
      TabIndex        =   6
      Top             =   1575
      Width           =   6000
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Para optimizar esta presentación de FLEXOFT es conveniente re-configurar su monitor."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   1155
      TabIndex        =   5
      Top             =   945
      Width           =   6000
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Atención: FLEXOFT ha detectado la configuración de su                     monitor."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   1155
      TabIndex        =   3
      Top             =   315
      Width           =   6000
   End
End
Attribute VB_Name = "WINDEMO2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Close
    End
End Sub

Private Sub Command2_Click()
    PrintForm
End Sub

Private Sub Command3_Click()
    WINDEMO1.Timer1.Enabled = True
    WINDEMO1.Timer2.Enabled = True
    WINDEMO2.Hide
End Sub

Private Sub Form_Load()
    WINDEMO1.Timer1.Enabled = False
    WINDEMO1.Timer2.Enabled = False
End Sub

