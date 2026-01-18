VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form DERELPRE 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Derechos de Acceso de Usuarios"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   8910
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
   ScaleHeight     =   5610
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   15
      Top             =   1320
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Nuevo Derecho"
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
      Left            =   4920
      TabIndex        =   14
      Top             =   1200
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.CommandButton Command3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   7005
      Picture         =   "DERELPRE.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   315
      Width           =   880
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Buscar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   4920
      TabIndex        =   10
      Top             =   210
      Width           =   2025
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
         Height          =   360
         Left            =   195
         TabIndex        =   11
         Text            =   " "
         ToolTipText     =   "<ENTER> para Búsqueda de Texto"
         Top             =   330
         Width           =   1590
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   3795
      Left            =   210
      ScaleHeight     =   3735
      ScaleWidth      =   8535
      TabIndex        =   7
      Top             =   1680
      Width           =   8600
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3435
         ItemData        =   "DERELPRE.frx":066A
         Left            =   -10
         List            =   "DERELPRE.frx":066C
         TabIndex        =   13
         Top             =   330
         Width           =   8500
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   10300
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción                                                                   L    P      Lista N°"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   105
            TabIndex        =   9
            Top             =   60
            Width           =   9870
         End
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         X1              =   0
         X2              =   6300
         Y1              =   3990
         Y2              =   3990
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Usuario FLEXOFT Seleccionado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1310
      Left            =   210
      TabIndex        =   3
      Top             =   210
      Width           =   4605
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   1005
         TabIndex        =   6
         Top             =   750
         Width           =   3480
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   240
         TabIndex        =   5
         Top             =   750
         Width           =   540
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " Número de Usuario:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   0
         TabIndex        =   4
         Top             =   240
         Width           =   1065
      End
   End
   Begin VB.CommandButton Command7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   7005
      Picture         =   "DERELPRE.frx":066E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   930
      Width           =   880
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   7920
      Picture         =   "DERELPRE.frx":0978
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   930
      Width           =   880
   End
   Begin VB.CommandButton command1 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   7920
      Picture         =   "DERELPRE.frx":0C82
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   315
      Width           =   880
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   750
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "DERELPRE.frx":0DCC
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   1590
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3600
      OleObjectBlob   =   "DERELPRE.frx":0E46
      Top             =   1200
   End
   Begin ACTIVESKINLibCtl.Skin Skin2 
      Left            =   3600
      OleObjectBlob   =   "DERELPRE.frx":107A
      Top             =   1200
   End
End
Attribute VB_Name = "DERELPRE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VACON%

Private Sub Command1_Click()
   Call CIERRARCH("*.*")
   Hide
End Sub

Private Sub Command2_Click()
   '
   
   Screen.MousePointer = 11
   '
   NUSER% = XVALO(Label2)
   CONDI$ = "USUARIO = " & NUSER%
   Call ACTUREGISTRO("DERELISTAPRE", "MARBORRA", CONDI$, 1, REGAF&, "NOMESS")
   USEL% = Val(Label2.Caption)
   For i& = 1 To List1.ListCount
      X$ = List1.List(i& - 1)
      DER1$ = UCase$(Mid$(X$, 45, 2))
      DER2$ = UCase$(Mid$(X$, 49, 2))
      CODAPLI$ = Mid$(X$, 53, 32)
      
      NULIST% = XVALO(Mid$(CODAPLI$, 11))
      If DER1$ = "SI" Then
         DERE1% = 1
         If DER2$ = "SI" Then
            DERE2% = 1
         End If
      End If
      DERECHO% = DERE1% + DERE2%
      If NULIST% > 0 And DERECHO% > 0 Then
10      RNDX$ = RANDSTRING$(-4)
        RND1% = XVALO(RNDX$)
        If CantRegistros("DERELISTAPRE", "Cod_Lista = " & RND1%, "NOMESS") > 0 Then GoTo 10
        Call CreaRegistro("DERELISTAPRE", "Cod_Lista", RND1%, "NOMESS")
        '
        CONDI$ = "COD_LISTA = " & RND1%
        Call ACTUREGISTRO("DERELISTAPRE", "Usuario", CONDI$, NUSER%, REGAF&, "NOMESS")
        Call ACTUREGISTRO("DERELISTAPRE", "DERECHO", CONDI$, DERECHO%, REGAF&, "NOMESS")
        Call ACTUREGISTRO("DERELISTAPRE", "MARBORRA", CONDI$, 0, REGAF&, "NOMESS")
        Call ACTUREGISTRO("DERELISTAPRE", "COD_LISTA", CONDI$, NULIST%, REGAF&, "NOMESS")
        DERE1% = 0: DERE2% = 0
      End If
      '
   Next i&
   '
   CONDI$ = "USUARIO = " & NUSER%
   CONDI2$ = CONDI$ + " and marborra = 1"
   Call BORRAREGISTROS("DERELISTAPRE", CONDI2$, REGAF&, "NOMESS")
   Screen.MousePointer = 1
   Hide
   '
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   OPPX% = COMALTER%("Seleccione el Formato del Listado:\\    Completo    \Solo Derechos Asignados")
   Call HEADERS("DERUSER", "")
   Call HEADERS("DERUSER", "Usuario: " + TRIM$(Label4) + " (" + TRIM$(Label2) + ")")
   '
   jj& = 0
   For UI% = 1 To List1.ListCount
     XXX$ = List1.List(UI% - 1)
        If OPPX% = 1 Then
           ZZZ$ = Space$(64)
           Call REPLA(ZZZ$, Mid$(XXX$, 53, 12), 1, 12)
           Call REPLA(ZZZ$, Mid$(XXX$, 1, 36), 13, 48)
           Call REPLA(ZZZ$, Mid$(XXX$, 45, 2), 61, 2)
           Call REPLA(ZZZ$, Mid$(XXX$, 49, 2), 63, 2)
           jj& = jj& + 1
           Call GRAREG("DERUSER", ZZZ$, jj&)
        ElseIf OPPX% = 2 Then
           If TRIM$(Mid$(XXX$, 45, 6)) <> "" Then
              ZZZ$ = Space$(64)
              Call REPLA(ZZZ$, Mid$(XXX$, 53, 12), 1, 12)
              Call REPLA(ZZZ$, Mid$(XXX$, 1, 36), 13, 48)
              Call REPLA(ZZZ$, Mid$(XXX$, 45, 2), 61, 2)
              Call REPLA(ZZZ$, Mid$(XXX$, 49, 2), 63, 2)
              jj& = jj& + 1
              Call GRAREG("DERUSER", ZZZ$, jj&)
           End If
        Else
            Exit Sub
        End If
   Next UI%
   Call SETULTREG("DERUSER", jj&)
   Call CIERRARCH("DERUSER")
   '
   Call CONECRUN("*DERUSER", "Listado Derechos de Acceso por Usuario")
   '
   Command3.Enabled = True
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
10 OBX$ = OBJPLA$("NUEVDERE", VDEF$)
   If OBX$ = "" Then
        GoTo 99
   End If
   
   CODI$ = UCase$(TRIM$(Mid$(OBX$, 1, 16)))
   Descr$ = UCase$(TRIM$(Mid$(OBX$, 17, 24)))
   '
   'VALIDO
      If TRIM$(CODI$) = "" Then
     Call COMUNI("Falta Codigo de Derecho")
     VDEF$ = OBX$: GoTo 10
   End If
   '
   If TRIM$(Descr$) = "" Then
     Call COMUNI("Falta Descripción de Derecho")
     VDEF$ = OBX$: GoTo 10
   End If
   '
   'PASO LOS PARAMETROS A LA FUNCION DERACCE
   If DERACCE%(CODI$, Descr$) > 0 Then
   End If
   '
   'ACTUALIZO EL LIST
   USEL% = XVALO(Label2)
   Call ACLISDER(USEL%, CODI$)
   
   Text1 = CODI$
   Text1.SetFocus
   
   
99 Command4.Enabled = True

End Sub




Private Sub Form_Load()
   Call APLICACIONSKINS(Me)
End Sub





Private Sub List1_Click()
   If VACON% = 1 Then
       Exit Sub
   End If
   ILISTA& = List1.ListIndex
   Z$ = List1.TEXT
   If Mid$(Z$, 49, 2) = "Si" Then
        Call REPLA(Z$, "       ", 45, 7)
      ElseIf Mid$(Z$, 45, 2) = "Si" Then
        Call REPLA(Z$, "Si", 49, 2)
      Else
        Call REPLA(Z$, "Si     ", 45, 7)
   End If
   List1.RemoveItem ILISTA&
   List1.AddItem Z$, ILISTA&
   VACON% = 1
   List1.ListIndex = ILISTA&
   VACON% = 0
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
'   If KeyCode = 46 Then
'      CODAP$ = TRIM$(Mid$(List1.TEXT, 53, 32))
'      CODAP12$ = TRIM$(Left$(CODAP$, 12))
'      JJ& = 0
'      For U& = 1 To ULTREG&("INDIAPLI")
'         XX$ = REGLEIDO$("INDIAPLI", U&)
'         If TRIM$(Left$(XX$, 12)) <> CODAP12$ Then
'           JJ& = JJ& + 1
'           Call GRAREG("INDIAPLI", XX$, JJ&)
'         End If
'      Next U&
'      Call SETULTREG("INDIAPLI", JJ&)
'      '
'      Call ABRETACON
'      SQLX$ = "DELETE FROM INDI_APLI WHERE COD_APLI = '" & CODAP$ & "'"
'      On Error Resume Next
'      BaDaConsul.Execute SQLX$
'      If Err Then
'          On Error GoTo 0
'          Exit Sub
'        Else
'          ILISTA& = List1.ListIndex
'          List1.RemoveItem ILISTA&
'      End If
'      On Error GoTo 0
'   End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   
   If KeyAscii% = 13 Then
     VACON% = 1
     INIBUS& = List1.ListIndex + 1
     TEBUS$ = TRIM$(UCase$(Text1.TEXT))
     For KU& = INIBUS& + 1 To List1.ListCount
       If InStr(UCase$(List1.List(KU& - 1)), TEBUS$) > 0 Then
         List1.ListIndex = KU& - 1
         VACON% = 0
         Exit Sub
       End If
     Next KU&
     For KU& = 1 To INIBUS&
       If InStr(UCase$(List1.List(KU& - 1)), TEBUS$) > 0 Then
         List1.ListIndex = KU& - 1
         VACON% = 0
         Exit Sub
       End If
     Next KU&
     VACON% = 0
   End If
End Sub
