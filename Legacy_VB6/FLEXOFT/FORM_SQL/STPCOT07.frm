VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form STPCOT07 
   BackColor       =   &H00E2D3C0&
   Caption         =   "Configuración C.O.T."
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   -450
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   5745
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00624E28&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6225
      Left            =   4920
      ScaleHeight     =   6225
      ScaleWidth      =   855
      TabIndex        =   2
      Top             =   -105
      Width           =   855
      Begin VB.CommandButton Command2 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   105
         Picture         =   "STPCOT07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Graba la Configuración "
         Top             =   1080
         Width           =   650
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   105
         Picture         =   "STPCOT07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Abandona la Aplicación"
         Top             =   315
         Width           =   650
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E2D3C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5355
      Left            =   50
      TabIndex        =   0
      Top             =   0
      Width           =   4755
      Begin VB.Frame Frame4 
         BackColor       =   &H00E2D3C0&
         Caption         =   "JURISDICCION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   18
         Top             =   4440
         Width           =   4455
         Begin VB.OptionButton Option2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "CABA"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2760
            TabIndex        =   20
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "BS AS"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   600
            TabIndex        =   19
            Top             =   360
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E2D3C0&
         Caption         =   "CIUDAD AUTONOMA DE BS AS"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   120
         TabIndex        =   10
         Top             =   1920
         Width           =   4455
         Begin VB.Timer tmrLecturaToken 
            Enabled         =   0   'False
            Interval        =   1000
            Left            =   1200
            Top             =   1920
         End
         Begin VB.CommandButton btnObtenerToken 
            Caption         =   "Obtener Token"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2760
            TabIndex        =   17
            Top             =   2040
            Width           =   1455
         End
         Begin VB.TextBox txtVencimientoTokenAGIP 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   1200
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   16
            Top             =   1560
            Width           =   3015
         End
         Begin VB.TextBox txtSignAGIP 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            Left            =   840
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   12
            Top             =   960
            Width           =   3375
         End
         Begin VB.TextBox txtTokenAGIP 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            Left            =   840
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   11
            Top             =   360
            Width           =   3375
         End
         Begin VB.Label Label5 
            BackColor       =   &H00E2D3C0&
            Caption         =   "Vencimiento"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   15
            Top             =   1680
            Width           =   1080
         End
         Begin VB.Label Label4 
            BackColor       =   &H00E2D3C0&
            Caption         =   "Sign"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   14
            Top             =   1300
            Width           =   1440
         End
         Begin VB.Label Label3 
            BackColor       =   &H00E2D3C0&
            Caption         =   "Token"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   13
            Top             =   720
            Width           =   1440
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E2D3C0&
         Caption         =   "BUENOS AIRES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   4455
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
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
            Left            =   1680
            TabIndex        =   7
            Top             =   360
            Width           =   2535
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
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
            Left            =   1680
            TabIndex        =   6
            Top             =   960
            Width           =   2535
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E2D3C0&
            Caption         =   "Password C.O.T."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   120
            TabIndex        =   9
            Top             =   500
            Width           =   1440
         End
         Begin VB.Label Label2 
            BackColor       =   &H00E2D3C0&
            Caption         =   "Código de Producto Genérico s/ARBA"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   120
            TabIndex        =   8
            Top             =   900
            Width           =   1440
         End
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   2625
         OleObjectBlob   =   "STPCOT07.frx":0454
         Top             =   0
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
         Height          =   225
         Index           =   0
         Left            =   120
         OleObjectBlob   =   "STPCOT07.frx":0688
         TabIndex        =   1
         Top             =   1320
         Visible         =   0   'False
         Width           =   1485
      End
   End
End
Attribute VB_Name = "STPCOT07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnObtenerToken_Click()

    btnObtenerToken.Enabled = False
    
    Call ObtenerTokenAGIPCOT
    
End Sub

Private Sub Command1_Click()
    '
    Command1.Enabled = False
    Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    Unload Me
    Command1.Enabled = True
    '
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   If Len(TRIM$(Text2.TEXT)) > 6 Then
      Call COMUNI("Código Genérico Longitud Máxima Permitida 6 Caracteres")
      GoTo 99
   End If
   Call GRACONTROL("REMITO", "PASSCOT", TRIM$(Me.Text1))
   Call GRACONTROL("REMITO", "CODIARBA", TRIM$(Me.Text2))
'   Call GRACONTROL("REMITO", "RUTACOT", TRIM$(Me.Label2))
99 Command2.Enabled = True
End Sub


Private Sub Form_Load()
    '
    Call APLICACIONSKINS(Me, Picture1)

    Text1 = Control$("REMITO", "PASSCOT")
    Text2 = Control("REMITO", "CODIARBA")
    '
    txtTokenAGIP = Control("REMELEC", "TOKEAGIP")
    txtSignAGIP = Control("REMELEC", "SIGNAGIP")
    txtVencimientoTokenAGIP = Control("REMELEC", "EXPIAGIP")
    '
    If Control("REMELEC", "MODOCOT") = "ARBA" Then Option1.Value = 1
    If Control("REMELEC", "MODOCOT") = "AGIP" Then Option2.Value = 1
    
End Sub

Sub ObtenerTokenAGIPCOT()

    Screen.MousePointer = 11
    
    ' VERIFICA EXISTENCIA DE PROGRAMA DE OBTENCION DEL TOKEN
    RutaPrograma$ = App.Path
    If EXISTE%(RutaPrograma$ & "\ClienteLoginCms_COT.exe") = 0 Then GoTo FaltaProgramaToken
    Call FileCopy(RutaPrograma$ & "\ClienteLoginCms_COT.exe", LUCOM$ & "COT\ClienteLoginCms_COT.exe")
    
    ' VERIFICA EXISTENCIA DEL CERTIFICADO DIGITAL
    If EXISTE%(LUDAT$ & "micertificadocot.pfx") = 0 Then GoTo FaltaCertificadoDigital
    Call FileCopy(LUDAT$ & "micertificadocot.pfx", LUCOM$ & "COT\micertificadocot.pfx")
            
    ' ELIMINA ULTIMO TOKEN OBTENIDO
    If EXISTE%(LUCOM$ & "COT\TOKEN_COT.TXT") = 1 Then Call Kill(LUCOM$ & "COT\TOKEN_COT.TXT")
            
    Resultado& = Shell(LUCOM$ & "COT\ClienteLoginCms_COT.exe", vbMinimizedFocus)
    If Resultado& < 0 Then GoTo FalloAlEjecutarProgramaToken
     
    tmrLecturaToken.Enabled = True
    
99  Screen.MousePointer = 1
    Exit Sub
    
FaltaProgramaToken:
    Call COMUNI("Imposible Obtener Token.\ \Falta Programa de Solicitud de Token.")
    GoTo 99

FaltaCertificadoDigital:
    Call COMUNI("Imposible Obtener Token.\ \Falta Certificado Digital.")
    GoTo 99
    
FalloAlEjecutarProgramaToken:
    Call COMUNI("Fallo al Ejecutar Programa de Solicitud de Token")
    GoTo 99

End Sub

Private Sub Option1_Click()

    Frame3.Enabled = False
    Text1.Enabled = True
    
    Call GRACONTROL("REMELEC", "MODOCOT", "ARBA")
    
End Sub

Private Sub Option2_Click()

    Frame3.Enabled = True
    Text1.Enabled = False
    Call GRACONTROL("REMELEC", "MODOCOT", "AGIP")

End Sub

Private Sub tmrLecturaToken_Timer()

    Static Intentos%
    
    Screen.MousePointer = 11
    
    If EXISTE%(LUCOM$ + "COT\TOKEN_COT.TXT") > 0 Then
        tmrLecturaToken.Enabled = False
        NARCHIVO% = FILEOPEN%(LUCOM$ + "COT\TOKEN_COT.TXT", 1, 128)  'ABRE EL ARCHIVO DE LECTURA
        Do While Not EOF(NARCHIVO%) ' RECORRE
            Line Input #NARCHIVO%, REGISTRO$ ' LEE CADA REGISTRO
            
            If InStr(1, REGISTRO$, "<Token>") > 0 Then
                FinEtiqueta% = InStr(1, REGISTRO$, "</Token>")
                Token$ = Mid$(REGISTRO$, 1, FinEtiqueta% - 1)
                Token$ = Mid$(Token$, Len("<Token>") + 1)
                txtTokenAGIP = Token$
            End If

            If InStr(1, REGISTRO$, "<Sign>") > 0 Then
                FinEtiqueta% = InStr(1, REGISTRO$, "</Sign>")
                Sign$ = Mid$(REGISTRO$, 1, FinEtiqueta% - 1)
                Sign$ = Mid$(Sign$, Len("<Sign>") + 1)
                txtSignAGIP = Sign$
            End If

            If InStr(1, REGISTRO$, "<Expiracion>") > 0 Then
                FinEtiqueta% = InStr(1, REGISTRO$, "</Expiracion>")
                Expira$ = Mid$(REGISTRO$, 1, FinEtiqueta% - 1)
                Expira$ = Mid$(Expira$, Len("<Expiracion>") + 1)
                txtVencimientoTokenAGIP = Expira$
            End If
        Loop
        Close #NARCHIVO%
        
        Call GRACONTROL("REMELEC", "TOKEAGIP", Token$)
        Call GRACONTROL("REMELEC", "SIGNAGIP", Sign$)
        Call GRACONTROL("REMELEC", "EXPIAGIP", Expira$)
        btnObtenerToken.Enabled = True
        Screen.MousePointer = 1
        Exit Sub
    End If
    
    Intentos% = Intentos% + 1
    If Intentos% > 60 Then
        tmrLecturaToken.Enabled = False
        btnObtenerToken.Enabled = True
        Screen.MousePointer = 1
    End If
      
End Sub
