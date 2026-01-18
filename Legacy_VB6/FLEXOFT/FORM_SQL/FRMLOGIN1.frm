VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FRMLOGIN1 
   Caption         =   "Usuario"
   ClientHeight    =   3315
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   5640
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   1800
      OleObjectBlob   =   "FRMLOGIN1.frx":0000
      Top             =   1440
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ingreso de Usuario y Clave"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      Begin VB.CommandButton Command1 
         Caption         =   "Ingresar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3840
         TabIndex        =   6
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1680
         Width           =   4215
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   960
         TabIndex        =   1
         Top             =   360
         Width           =   4215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   495
         Left            =   40
         OleObjectBlob   =   "FRMLOGIN1.frx":0234
         TabIndex        =   3
         Top             =   480
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   495
         Left            =   45
         OleObjectBlob   =   "FRMLOGIN1.frx":02A2
         TabIndex        =   4
         Top             =   1800
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   375
         Left            =   960
         OleObjectBlob   =   "FRMLOGIN1.frx":030C
         TabIndex        =   5
         Top             =   1080
         Width           =   4215
      End
   End
End
Attribute VB_Name = "FRMLOGIN1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Private mcIni     As clsIni               ' guardar los datos en un archivo .ini
 Dim INI_PATH      As String

Private Sub Command1_Click()
    

    'SI NO HAY NINGUN REGISTRO ENTRA SIN VALIDAR
  SQLX$ = "SELECT COUNT (*) AS CANTI FROM USUARIOS"
  Set RS1 = BASSSEX.Execute(SQLX$)

  CANTEG& = Val(RS1!CANTI)

     If CANTEG& < 1 Then
'        ValidaUsuario% = 1
        NUMEUSER% = 1
        NOMBREUSUARIO$ = "SUPERVISOR"
        SOYSUPERVISOR% = 1
        mdiprincipal.Show
        Unload Me
        Exit Sub
     End If

   
   Call ValidaUsuario
   
   If ValidaClave% > 0 Then
     CONDI$ = "USUARIO = '" & UCase$(TRIM$(Text1)) & "' AND CLAVE = '" & UCase$(TRIM$(Text2)) & "'"
     NUMEUSER% = XVALO(VALOADMIN("USUARIOS", "NUMEUSUARIO", CONDI$, "NOMESS"))
     NOMBREUSUARIO$ = VALOADMIN("USUARIOS", "NOMBRE", CONDI$, "NOMESS")
     SOYSUPERVISOR% = XVALO(VALOADMIN("USUARIOS", "DERE_SUPERVISOR", CONDI$, "NOMESS"))
     mdiprincipal.Show
     Unload Me
   End If
End Sub

Private Sub Command2_Click()
  
End Sub
Private Sub Form_Initialize()

If App.PrevInstance = True Then
    
    
    MsgBox " La Aplicación Yá Se Encuentra Cargada ", vbInformation

    '' Finaliza el programa
    End

End If

End Sub


Private Sub Form_Load()

    Call APLICAR_SKIN1(Me, App.Path & "\SKINS\dogmax.skn")

'    If VALIDACONTRASEÑA% > 0 Then
'         IngClave.Show 1
'    End If

''    Call CREADIRECTORIO
''    POS% = InStrRev(App.Path, "\")
''    LUDAT$ = Mid$(App.Path, 1, POS% - 1) + "\DATOS\"
''    LUBAS$ = App.Path + "\"
''    LUCOM$ = "C:\PROGRAMACION_INDEPENDIENTE\"
''
''    Call ABRECONEXION
''    POS% = InStrRev(App.Path, "\")
''    LUDAT$ = Mid$(App.Path, 1, POS% - 1) + "\DATOS\"

'   Call VERJOBID(OKEY%)
'   If OKEY% < 1 Then Call FINAL("")
   '
'   Call CONAPLI(APLINVO$, TERMINA%) ' Busca la aplicacion a mostrar
'   If TERMINA% > 0 Then Call FINAL("")
   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   '
'   UTILIZA_SKIN% = 1
'   AMASTOK07.Show
'   Hide

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      
      If ValidaUsuario% > 0 Then
        Text2.SetFocus
      Else
        Text1.SetFocus
      End If
   End If

End Sub

Function ValidaUsuario%()
    '
     '
     Static INGRESO%
     CONDI$ = "USUARIO = '" & UCase$(TRIM$(Text1)) & "'"
     '
     ValidaUsuario% = 0
     'VALIDO SI EXISTE
     CNRG& = CantRegistros&("USUARIOS", CONDI$)
     If CNRG& < 1 Then
       MsgBox " USUARIO INVÁLIDO "
       Text1.SetFocus
       INGRESO% = INGRESO% + 1
       If INGRESO% > 6 Then End: Close
     Else
       ValidaUsuario% = 1
       nombre$ = VALOADMIN("USUARIOS", "NOMBRE", CONDI$, "NOMESS")
       APELL$ = VALOADMIN("USUARIOS", "APELLIDO", CONDI$, "NOMESS")
       SkinLabel3.Caption = nombre$ + " " + APELL$
     End If

     
End Function
Function ValidaClave%()

     'SI NO TIENE CLAVE REGISTRADA INGRESA SOLO CON EL USUARIO

     Static INGRESOCLAVE%
     ValidaClave% = 0
     '
     If TRIM$(SkinLabel3.Caption) = "" Then
       MsgBox "FALTA INGRESAR USUARIO O EL MISMO NO ES VALIDO"
       Exit Function
     End If
     '
     CONDI$ = "CLAVE = '" & UCase$(TRIM$(Text2)) & "' and USUARIO = '" & UCase$(TRIM$(Text1)) & "'"
     '
     'VALIDO SI EXISTE
     CNRG& = CantRegistros&("USUARIOS", CONDI$)
     If CNRG& < 1 Then
       MsgBox " CLAVE INVÁLIDA "
       INGRESOCLAVE% = INGRESOCLAVE% + 1
       If INGRESOCLAVE% > 6 Then End: Close
     Else
       ValidaClave% = 1
     End If
End Function

Private Sub Text2_KeyPress(KeyAscii As Integer)

  If KeyAscii = 13 Then
    Call Command1_Click
  End If
End Sub
Function VALIDACONTRASEÑA%()

''       'OJO PARA SE DEBE DECLARAR ESTAS DOS EN GENERAL DEL FRM QUE SE ESTA USANDO
''       Private mcIni     As clsIni               ' guardar los datos en un archivo .ini
''       Dim INI_PATH      As String
        RESPUESTA% = 0
        Set mcIni = New clsIni
        INI_PATH = App.Path & "\config.ini"
       
        With mcIni
            'CLAVE$ = .Encriptar(App.EXEName, .getValue(INI_PATH, "datos", "CLAVEIDA"), 2)
            DISLOC$ = DISCOLOCAL$ ' LEE EL DISCO LOCAL
            'COMPARA CON EL GUARDADO EN EL ARVHIVO INI
            CONTRACLAVE$ = .Encriptar(App.EXEName, .getValue(INI_PATH, "datos", "CLAVEVUELTA"), 2)
            CLAVEFUN$ = CONTRACLA$(DISLOC$)
            If CONTRACLAVE$ <> CLAVEFUN$ Then
               RESPUESTA% = 1
            End If
       End With
       '
       VALIDACONTRASEÑA% = RESPUESTA%
End Function

