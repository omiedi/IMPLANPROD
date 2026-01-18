VERSION 5.00
Begin VB.Form LOGVEND 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asignación de Vendedor a Usuario"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4515
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDC0&
      Caption         =   "Configuración de Vendedor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2700
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   3585
      Begin VB.TextBox Text2 
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
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Doble Click Para Limpiar Casillero"
         Top             =   2160
         Width           =   3135
      End
      Begin VB.TextBox TXTLABEL 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Index           =   15
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   "Vendedor"
         Top             =   1680
         Width           =   900
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1680
         Width           =   400
      End
      Begin VB.CommandButton Command22 
         Caption         =   "."
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
         Left            =   1140
         TabIndex        =   11
         Top             =   1680
         Width           =   175
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         IMEMode         =   3  'DISABLE
         Left            =   2805
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   390
         Width           =   570
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   240
         TabIndex        =   10
         Top             =   1005
         Width           =   3165
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario N°:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   9
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00404040&
      Height          =   2730
      Left            =   3600
      ScaleHeight     =   2670
      ScaleWidth      =   1050
      TabIndex        =   4
      Top             =   0
      Width           =   1110
      Begin VB.CommandButton Command2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   105
         Picture         =   "LOGVEND.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1470
         Width           =   645
      End
      Begin VB.CommandButton command1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   105
         Picture         =   "LOGVEND.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   105
         Width           =   645
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -300
         Picture         =   "LOGVEND.frx":0454
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   2040
         Width           =   1515
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Presentación de Costo y Precios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   210
      TabIndex        =   0
      Top             =   7770
      Width           =   5160
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cos.Repos. Calculado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1920
         TabIndex        =   3
         Top             =   360
         Width           =   1305
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Precio Base (C.Rep. + 15 %)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3360
         TabIndex        =   2
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cos.Repos. Informado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   1905
      End
   End
End
Attribute VB_Name = "LOGVEND"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public YACAR%
Dim FF0 As String * 128
Dim cmd$

Dim CLA$, CLB$, CLAV$
Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command2_Click()
    '
    Command2.Enabled = False
    USX% = XVALO(TEXT1)
    VEX% = XVALO(Text4)
    If USX% < 1 Then
      Call COMUNI("Falta N° de Usuario")
      GoTo 99
    End If
    If VEX% < 1 Then
      Call COMUNI("Falta N° de Vendedor")
      GoTo 99
    End If

    Call GRACONTROL("LOGVEND", SAFETEXT(USX%), SAFETEXT$(VEX%))
    
    
    Call CIERRARCH("*.*")
    Call COMUNI("Re-Configuración Completada.\Re-Inicie Ahora la Aplicación '" + UCase$(App.EXEName) + "'.")
    Call FINAL("")
    '
99  Command2.Enabled = True

End Sub

Function VendeVsUsuario%(NUS%)
   VendeVsUsuario% = XVALO(CONTROL("LOGVEND", SAFETEXT(NUS%)))
End Function

Private Sub Command3_Click()
'    Call GRACONTROL("LOGVENDE", "VENDELOG", XVALO(Text4))
End Sub



Private Sub Command22_Click()
    Call SELECHO("VENDEDOR")
    VDEVU$ = TRIM$(VALACT1$("VENDEDOR"))
    If VDEVU$ <> "" Then
       Text4.TEXT = VDEVU$
    End If
End Sub

Private Sub Form_Load()
    US% = USNBR% Mod 100
    Text4.TEXT = VendeVsUsuario%(US%)
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
'    If KeyAscii <> 13 Then Exit Sub
'    VUSER$ = LOADFILE$("USER.DAT")
'    USNBR% = 0
'    USERNA$ = ""
'    CONTRA$ = ""
'    '
'    CAUS% = Len(VUSER$) / 64
''    If CAUS% < 1 Then GoTo 31
'    MNGR% = 0
'    For JJ% = 1 To CAUS%
'      F0$ = Mid$(VUSER$, 64 * JJ% - 63, 64)
'      Ident$ = Mid$(F0$, 31, 12)
'        If Left$(Ident$, 1) = "@" Then Ident$ = DESENCRI$(Mid$(F0$, 32, 11)) + " "
'        If LTrim$(RTrim$(Ident$)) = "" Then GoTo 19
'      DENO$ = Left$(F0$, 30)
'        If Left$(DENO$, 1) = "@" Then DENO$ = DESENCRI$(Mid$(F0$, 2, 29)) + " "
'      CLAV$ = Mid$(F0$, 43, 12)
'        If Left$(CLAV$, 1) = "@" Then CLAV$ = DESENCRI$(Mid$(F0$, 44, 11)) + " "
'        CLAV$ = REPLACAR$(CLAV$, Chr$(0), Chr$(32))
'      CLAV$ = TRIM$(CLAV$) + Space$(12)
'      If TRIM$(UCase$(Ident$)) = TRIM$(UCase$(Text1)) Then
'         Label1 = DENO$
'         Text3.Enabled = True
'         Text3.SetFocus
'         Exit Sub
'      End If
'
'19  Next JJ%
'    '
'40  MsgBox "Usuario no Definido"
'    CAIN% = CAIN% + 1
'    If CAIN% > 3 Then
'        MsgBox "Superado Limite Intentos de Acceso !!!"
'        Close: End
'    End If
'    Text1.TEXT = ""
'    CLA$ = "": CLB$ = "": CLAVE$ = ""
'    Text1.SetFocus
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
'   Static CAIN%
'   If KeyAscii <> 13 Then Exit Sub
'
'     CONTRA$ = Left$(Text3 + Space$(12), 12)
'     If TRIM$(Left$(CLAV$, 12)) = TRIM$(Left$(CONTRA$, 12)) Then
'
'
'     Else
'        MsgBox "Usuario no Definido"
'        Text3.SetFocus
'     End If
'




End Sub

Private Sub Text4_Change()
  Text2 = ECOARCH$("VENDEDOR", Chr$(XVALO(Text4.TEXT)))


End Sub
Sub BUSER()
    '
    Static CAIN%
    CLA$ = Left$(UCase$(CLAVE$) + Space$(12), 12)
    '
    If TRIM$(CLA$) = "???" Then
      CLAX$ = RANDSTRING$(-6)
      CONTRASE$ = InputBox$("Ingrese Contraseña", "Santo y Seña (" + CLAX$ + ")")
      If CONTRASE$ <> CONTRACLA$(CLAX$) Then GoTo 40
      '
      HOII% = HOY(HOS$): FLAB% = HOII% + 33
      DEF$ = Left$(LOADFILE$("FLXCTRL.EXE"), 129)
      If Len(DEF$) < 129 Then GoTo 20
      Call REPLA(DEF$, MKI$(FLAB%), 37, 2)
      Call REPLA(DEF$, MKI$(0), 39, 2)
      DIVERI% = 0: For UI% = 1 To 127
                     DIVERI% = DIVERI% + Asc(Mid$(DEF$, UI%, 1)) - 123
                   Next UI%
                   Call REPLA(DEF$, MKI$(DIVERI%), 128, 2)
      ABC$ = LOADFILE$("FLXCTRL.EXE")
      Call REPLA(ABC$, DEF$, 1, 129)
      Call SAVEFILE("FLXCTRL.EXE", ABC$)
      If ABC$ <> LOADFILE$("FLXCTRL.EXE") Then
        MsgBox "Acceso Denegado a Controlador"
      End If
      '
20    QWERTY% = 101: USNBR% = 101
      USERNA$ = "Supervisor General"
      'Call GRADATRAM
      Exit Sub
    End If
    '
30  If EXISTE%("USER.DAT") < 1 Then
31      QWERTY% = 101
        USNBR% = 101
        USERNA$ = "Supervisor General"
        'Call GRADATRAM
        Exit Sub
    End If
    '
    VUSER$ = LOADFILE$("USER.DAT")
    USNBR% = 0
    USERNA$ = ""
    CONTRA$ = ""
    '
    CAUS% = Len(VUSER$) / 64
    If CAUS% < 1 Then GoTo 31
    MNGR% = 0
    For JJ% = 1 To CAUS%
      F0$ = Mid$(VUSER$, 64 * JJ% - 63, 64)
      Ident$ = Mid$(F0$, 31, 12)
        If Left$(Ident$, 1) = "@" Then Ident$ = DESENCRI$(Mid$(F0$, 32, 11)) + " "
        If LTrim$(RTrim$(Ident$)) = "" Then GoTo 19
      DENO$ = Left$(F0$, 30)
        If Left$(DENO$, 1) = "@" Then DENO$ = DESENCRI$(Mid$(F0$, 2, 29)) + " "
      CLAV$ = Mid$(F0$, 43, 12)
        If Left$(CLAV$, 1) = "@" Then CLAV$ = DESENCRI$(Mid$(F0$, 44, 11)) + " "
        CLAV$ = REPLACAR$(CLAV$, Chr$(0), Chr$(32))
      CLAV$ = TRIM$(CLAV$) + Space$(12)
      If CLA$ = Ident$ Then
        NUSER% = JJ%
        Label1 = DENO$
        USERNA$ = TRIM$(DENO$)
        Text3.Enabled = True
        CONTRA$ = ""
        Text3.SetFocus
        Exit Sub
      End If
19  Next JJ%
    '
40  MsgBox "Usuario no Definido"
    CAIN% = CAIN% + 1
    If CAIN% > 3 Then
        MsgBox "Superado Limite Intentos de Acceso !!!"
        Close: End
    End If
    TEXT1.TEXT = ""
    CLA$ = "": CLB$ = "": CLAVE$ = ""
    TEXT1.SetFocus
    '
End Sub
Sub BUCLA()
   '
   Static CAIN%
   '
   If Left$(CLAV$, 12) = Left$(CONTRA$, 12) Then
       USNBR% = NUSER%
       If USNBR% = 1 Then
          USNBR% = 101
          If TRIM$(USERNA$) = "" Then
             USERNA$ = "Supervisor General"
          End If
       End If

       'Call GRADATRAM
      

     Else
       CAIN% = CAIN% + 1
       If CAIN% > 3 Then
          MsgBox "Superado Limite Permitido Intentos de Acceso !!!"
          Close: End
       End If
       '
       MsgBox "Contraseña Incorrecta"
       Text3 = ""
       Text3.SetFocus
   End If
   '
End Sub
