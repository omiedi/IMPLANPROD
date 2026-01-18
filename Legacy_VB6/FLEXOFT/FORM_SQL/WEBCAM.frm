VERSION 5.00
Begin VB.Form WEBCAM 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Captura Imagen desde WebCam"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   11040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   6480
      TabIndex        =   11
      Top             =   7560
      Width           =   3375
   End
   Begin VB.CommandButton Command8 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9840
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command7 
      Caption         =   "resolucion"
      Height          =   255
      Left            =   6000
      TabIndex        =   9
      Top             =   8160
      Width           =   1095
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7230
      Left            =   120
      ScaleHeight     =   7200
      ScaleWidth      =   9600
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   9630
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   7560
      Width           =   3375
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Guarda Imagen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9840
      TabIndex        =   6
      Top             =   6840
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Captura Imagen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9840
      TabIndex        =   5
      Top             =   6000
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7230
      Left            =   120
      ScaleHeight     =   480
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   4
      Top             =   120
      Width           =   9630
   End
   Begin VB.CommandButton Command4 
      Caption         =   "configurar"
      Height          =   255
      Left            =   4680
      TabIndex        =   3
      Top             =   8160
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "formato"
      Height          =   255
      Left            =   3360
      TabIndex        =   2
      Top             =   8160
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "detener"
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      Top             =   8160
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "iniciar"
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   8160
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Ruta"
      Height          =   255
      Left            =   5400
      TabIndex        =   13
      Top             =   7680
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Archivo"
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   7680
      Width           =   855
   End
End
Attribute VB_Name = "WEBCAM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oWebCam As ePerifericos
Dim nWebCam As nPerifericos
  
' botón que inicia la captura
'''''''''''''''''''''''''''''''''''''''
Private Sub Command1_Click()

    Call nWebCam.WebCam_Inicia(oWebCam)
  
End Sub
  
' botón para detener la captura
'''''''''''''''''''''''''''''''''''''''
Private Sub Command2_Click()
      
    Call nWebCam.WebCam_Desconecta(oWebCam)
  
End Sub
  
' Botón que abre el dialogo de formato
''''''''''''''''''''''''''''''''''''''''''''
Private Sub Command3_Click()

    Call nWebCam.WebCam_Formato(oWebCam)
    
End Sub

' Mostrar dialogo de Configuracion de la WebCam
''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Command4_Click()

    Call nWebCam.WebCam_Configuracion(oWebCam)

End Sub
  
' CAPTURA IMAGEN
Private Sub Command5_Click()

    Call nWebCam.WebCam_CapturaImagen(oWebCam)
    Picture2.Visible = True
    Command8.Visible = True

End Sub

' GUARDA CAPTURA

Private Sub Command6_Click()

    If Picture2.Visible Then
        oWebCam.WEBCAM.RutaGuardaCaptura = Text2
        oWebCam.WEBCAM.NombreArchivoCaptura = Text1
        
        Call nWebCam.WebCam_GuardaCaptura(oWebCam)
        Unload Me
    Else
        Call COMUNI("Falta capturar imagen")
    End If

End Sub
 
Private Sub Command7_Click()

    Call nWebCam.WebCam_Resolucion(oWebCam)

End Sub

Private Sub Command8_Click()

    Picture2.Picture = LoadPicture("")
    Picture2.Visible = False
    Command8.Visible = False
    
End Sub

Private Sub Form_Load()

    On Error GoTo Err_FalloConfiguracionCapturaFoto
    
    Set oWebCam = New ePerifericos
    Set oWebCam.WEBCAM.ObjetoReproduce = Picture1
    Set oWebCam.WEBCAM.ObjetoCaptura = Picture2
    
    Set nWebCam = New nPerifericos
    
    Call nWebCam.WebCam_Inicia(oWebCam)
    Call nWebCam.WebCam_Resolucion(oWebCam)
    
    On Error GoTo 0
    
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
    
99  Exit Sub
    

Err_FalloConfiguracionCapturaFoto:

    Unload Me

End Sub

Private Sub Form_Resize()

    'On Error Resume Next
    'Move (Screen.Width - Width) \ 29, (Screen.Height - Height) \ 29
    
End Sub
  
Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
    
    Call nWebCam.WebCam_Desconecta(oWebCam)
    
    Set oWebCam = Nothing
    Set nWebCam = Nothing
    
    On Error GoTo 0
    
End Sub

