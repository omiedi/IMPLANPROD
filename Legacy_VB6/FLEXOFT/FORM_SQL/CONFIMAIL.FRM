VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form CONFIMAIL 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   9225
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cmd1 
      Left            =   4680
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   2880
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   6
      Top             =   6720
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos de cuenta de Correo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7095
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   8160
      Begin VB.CommandButton Command3 
         Caption         =   "Ver Clave"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5115
         Picture         =   "CONFIMAIL.frx":0000
         TabIndex        =   26
         Top             =   1920
         Width           =   2820
      End
      Begin VB.CommandButton Command2 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1440
         Picture         =   "CONFIMAIL.frx":030A
         TabIndex        =   25
         Top             =   2520
         Width           =   175
      End
      Begin VB.TextBox TXTRUTAFIRMA 
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
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1680
         TabIndex        =   23
         Top             =   3000
         Width           =   6255
      End
      Begin VB.TextBox TXTRUTAPLANTILLAHTML 
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
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1680
         TabIndex        =   21
         Top             =   2520
         Width           =   6255
      End
      Begin VB.CommandButton Command9 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1440
         Picture         =   "CONFIMAIL.frx":0614
         TabIndex        =   19
         Top             =   3000
         Width           =   175
      End
      Begin VB.PictureBox Picture1 
         Height          =   2200
         Left            =   1800
         ScaleHeight     =   2145
         ScaleWidth      =   5940
         TabIndex        =   18
         Top             =   3720
         Width           =   6000
         Begin VB.Image Image1 
            Height          =   2200
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   6000
         End
      End
      Begin VB.TextBox TXTPUERTO 
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
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1440
         TabIndex        =   17
         Top             =   2040
         Width           =   1815
      End
      Begin VB.TextBox TXTSERVSMTP 
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
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1440
         TabIndex        =   15
         Top             =   880
         Width           =   6495
      End
      Begin VB.TextBox TXTCTAMAIL 
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
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1440
         TabIndex        =   14
         Top             =   360
         Width           =   6495
      End
      Begin VB.TextBox TXTCONTRASEÑA 
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
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1440
         PasswordChar    =   "*"
         TabIndex        =   13
         Top             =   1440
         Width           =   6495
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Test Mail"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4680
         TabIndex        =   4
         Top             =   6480
         Width           =   3180
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ruta de Imagen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   24
         Top             =   3120
         Width           =   1140
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ruta de Plantilla Html"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         TabIndex        =   22
         Top             =   2640
         Width           =   1140
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Imagen - Firma Pie de Página del Mail"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   120
         TabIndex        =   20
         Top             =   3720
         Width           =   1260
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Puerto:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   16
         Top             =   2160
         Width           =   1140
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Contraseña"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Top             =   1560
         Width           =   1140
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Servidor smtp (smtp.gmail.com) "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   1260
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta de Mail"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -240
         TabIndex        =   3
         Top             =   480
         Width           =   1500
      End
   End
   Begin VB.PictureBox Picture14 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   7425
      Left            =   8280
      ScaleHeight     =   7395
      ScaleWidth      =   1350
      TabIndex        =   0
      Top             =   -120
      Width           =   1380
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   120
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   9
         Top             =   6120
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -120
            TabIndex        =   10
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   120
         Picture         =   "CONFIMAIL.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   6480
         Width           =   690
      End
      Begin VB.CommandButton Command11 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "CONFIMAIL.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Salir"
         Top             =   225
         Width           =   690
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "CONFIMAIL.frx":0D72
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "CONFIMAIL.frx":0FA6
      TabIndex        =   7
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "CONFIMAIL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Command1.Enabled = False
   If TRIM$(TXTCTAMAIL) = "" Then
      Call COMUNI("Mail Inválido"):
      Call PINTATEXT(TXTCTAMAIL)
      GoTo 99
   End If
   If InStr(TXTCTAMAIL, "@") < 1 Then
      Call COMUNI("Mail Inválido"):
      Call PINTATEXT(TXTCTAMAIL)
      GoTo 99
   End If
   If TRIM$(TXTCONTRASEÑA) = "" Then
      Call COMUNI("Falta Contraseña"):
      Call PINTATEXT(TXTCONTRASEÑA)
      GoTo 99
   End If
   If TRIM$(TXTPUERTO) = "" Then
      Call COMUNI("Falta Puerto"):
      Call PINTATEXT(TXTPUERTO)
      GoTo 99
   End If
   If TRIM$(TXTSERVSMTP) = "" Then
      Call COMUNI("Falta Servidor Smtp"):
      Call PINTATEXT(TXTSERVSMTP)
      GoTo 99
   End If
   
   Usertex$ = SAFETEXT$(USNBR% Mod 100)
  
   Call GRACONTROL(Usertex$, "CTADMAIL", Me.TXTCTAMAIL)
   Call GRACONTROL(Usertex$, "CONTRASE", Me.TXTCONTRASEÑA)
   Call GRACONTROL(Usertex$, "PUERTOCO", Me.TXTPUERTO)
   Call GRACONTROL(Usertex$, "SERVERSM", Me.TXTSERVSMTP)
   Call GRACONTROL(Usertex$, "PLANTHTM", TXTRUTAPLANTILLAHTML)
   Call GRACONTROL(Usertex$, "FIRMAMAI", TXTRUTAFIRMA)
   
99 Command1.Enabled = True
End Sub

Private Sub Command11_Click()
   Unload Me
End Sub

Private Sub Command2_Click()
   RUTAX$ = SELECCION_Archivo$("TXT", CMD1)
   If RUTAX$ <> "" Then
      Call GRACONTROL("MAILNEW", "PLANMAIL", RUTAX$)
   End If
   Me.TXTRUTAPLANTILLAHTML = RUTAX$
   Screen.MousePointer = 1
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   USN% = USNBR% Mod 100
   If USN% <> 1 Then
      GoTo 99
   Else
      If Command3.Caption = "Ver Clave" Then
         TXTCONTRASEÑA.PasswordChar = ""
         Command3.Caption = "Ocultar Clave"
      Else
        TXTCONTRASEÑA.PasswordChar = "*"
        Command3.Caption = "Ver Clave"
      End If
   End If
   
99 Command3.Enabled = True
End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
   Call Command1_Click  'FUERZO LA GRABACION
    '
       PLANTILLAHTML$ = Control("MAILNEW", "PLANMAIL")
    
       '*********************************************************************************************
       'CREO EL ARCHIVO CON LOS DATOS DE ENVIO DE MAIL
       '*********************************************************************************************
       For I1% = 1 To 32767
         ARCHDAMAIL$ = "C:\Flexoft\" & "DATOSMAIL" & TRIM$(Str$(I1%)) & ".TXT"
         If EXISTE%(ARCHDAMAIL$) < 1 Then
           NAME_ARCH_DATMAIL$ = "DATOSMAIL" & TRIM$(Str$(I1%)) & ".TXT"
           NARCHI0% = FILEOPEN%(ARCHDAMAIL$, 2, 128)  ' ABRE EL ARCHIVO DE ESCRITURA
           Exit For
         End If
       Next I1%
       Usertex$ = SAFETEXT$(USNBR% Mod 100)
       CUENTAPROPIA1 = Control(Usertex$, "CTADMAIL")
       CONTRASEÑA1 = Control(Usertex$, "CONTRASE")
       PUERTO1 = Control(Usertex$, "PUERTOCO")
       SERVIDORSMTP1 = Control(Usertex$, "SERVERSM")
       
       Print #NARCHI0%, "|MAILRECEPTOR|" & CUENTAPROPIA1 'para probar se lo manda a la  misma cuenta
       Print #NARCHI0%, "|ASUNTO|      " & "MAIL DE PRUEBA"
       Print #NARCHI0%, "|CUENTAPROPIA|" & CUENTAPROPIA1
       Print #NARCHI0%, "|CONTRASEÑA|  " & CONTRASEÑA1
       Print #NARCHI0%, "|PUERTO|      " & PUERTO1
       Print #NARCHI0%, "|SERVIDORSMTP|" & SERVIDORSMTP1

       Close #NARCHI0%
       
       For I1% = 1 To 32767
         ARCH$ = "C:\Flexoft\" & "PLANTILLA_STANDARD" & TRIM$(Str$(I1%)) & ".TXT"
         If EXISTE%(ARCH$) < 1 Then
           NAME_ARCH_HTML$ = "PLANTILLA_STANDARD" & TRIM$(Str$(I1%)) & ".TXT"
           NARCHI2% = FILEOPEN%(ARCH$, 2, 128)  ' ABRE EL ARCHIVO DE ESCRITURA
            Print #NARCHI2%, "<HTML>"
            Print #NARCHI2%, "<HEAD>"
            Print #NARCHI2%, "</HEAD>"
            Print #NARCHI2%, "<H1><CENTER>ENCABEZADO1</CENTER></H1>"
            Print #NARCHI2%, "<H2><CENTER>ENCABEZADO2</CENTER></H2>"
            Print #NARCHI2%, "<HR>"
            Print #NARCHI2%, "<HR>"
            Print #NARCHI2%, "|CARGAR_LA_IMAGEN_FIRMA|"
            Print #NARCHI2%, "</BODY>"
            Print #NARCHI2%, "</HTML>"
            Close #NARCHI2%
            Exit For
         End If
       Next I1%

       '*********************************************************************************************
       'FIN EL ARCHIVO CON LOS DATOS DE ENVIO DE MAIL
       '*********************************************************************************************
       '
       ruta_archi_imag$ = Control(Usertex$, "FIRMAMAI") 'si esta configurado en el control para este usuario
       If TRIM$(ruta_archi_imag$) <> "" Then
          If EXISTE%(TRIM$(ruta_archi_imag$)) > 0 Then
             RUTA_ARCH_FIRMA$ = "C:\Flexoft\" & "FIRMA" & ".TXT"
             NARCHI3% = FILEOPEN%(RUTA_ARCH_FIRMA$, 2, 128) 'ABRE EL ARCHIVO DE LECTURA
             Print #NARCHI3%, ruta_archi_imag$
             Close #NARCHI3%
           End If
       End If
       '
       '
       
       XXX = Shell(App.Path + "\" & "EnvioMail2.exe " & NAME_ARCH_DATMAIL$ & " " & NAME_ARCH_HTML$ & " " & "FIRMA.TXT", 4)
99

   Command5.Enabled = True
End Sub

Private Sub Command9_Click()
   RUTAX$ = SELECCION_Archivo$("jpg", CMD1)
   If RUTAX$ <> "" Then
      Call GRACONTROL("MAILNEW", "FIRMAMAI", RUTAX$)
   End If
   TXTRUTAFIRMA = RUTAX$
   On Error Resume Next
   Image1.Picture = LoadPicture(RUTAX$)
   If Err > 0 Then MsgBox Err.Description
   Screen.MousePointer = 1
End Sub

Private Sub Form_Load()
   USN% = USNBR% Mod 100
   
   Me.Caption = "Datos de cuenta de Correo Usuario: " & USERNAM$(USN%)
   Usertex$ = SAFETEXT$(USNBR% Mod 100)
   TXTCTAMAIL = Control(Usertex$, "CTADMAIL")
   TXTCONTRASEÑA = Control(Usertex$, "CONTRASE")
   TXTPUERTO = Control(Usertex$, "PUERTOCO")
   TXTSERVSMTP = Control(Usertex$, "SERVERSM")
   TXTRUTAPLANTILLAHTML = Control(Usertex$, "PLANTHTM")
   TXTRUTAFIRMA = Control(Usertex$, "FIRMAMAI")
   On Error Resume Next
   Image1.Picture = LoadPicture(TXTRUTAFIRMA)
   On Error GoTo 0
End Sub

Private Sub TXTRUTAFIRMA_Change()
   If EXISTE%(Control(Usertex$, "FIRMAMAI")) > 0 Then
   On Error Resume Next
   Image1.Picture = LoadPicture(TXTRUTAFIRMA)
   On Error GoTo 0
   End If
End Sub
