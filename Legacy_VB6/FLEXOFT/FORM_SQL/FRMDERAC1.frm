VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FRMDERAC1 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Derechos de Acceso de Usuarios"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9885
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
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   9885
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command5 
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
      Height          =   375
      Left            =   2150
      TabIndex        =   17
      Top             =   800
      Width           =   255
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   5280
      OleObjectBlob   =   "FRMDERAC1.frx":0000
      Top             =   0
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      ItemData        =   "FRMDERAC1.frx":0234
      Left            =   240
      List            =   "FRMDERAC1.frx":0236
      TabIndex        =   10
      Top             =   3840
      Visible         =   0   'False
      Width           =   9450
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
      Left            =   5640
      TabIndex        =   9
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
      Left            =   7800
      Picture         =   "FRMDERAC1.frx":0238
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   270
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
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
      Left            =   5640
      TabIndex        =   5
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
         TabIndex        =   6
         ToolTipText     =   "<ENTER> para Búsqueda de Texto"
         Top             =   330
         Width           =   1590
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   3795
      Left            =   210
      ScaleHeight     =   3735
      ScaleWidth      =   9435
      TabIndex        =   4
      Top             =   1680
      Width           =   9495
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   375
         Left            =   0
         OleObjectBlob   =   "FRMDERAC1.frx":08A2
         TabIndex        =   16
         Top             =   0
         Width           =   10575
      End
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
         ItemData        =   "FRMDERAC1.frx":09C4
         Left            =   0
         List            =   "FRMDERAC1.frx":09C6
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   360
         Width           =   9450
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
      BackColor       =   &H00C0FFC0&
      Caption         =   "Usuario Seleccionado"
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
      Left            =   240
      TabIndex        =   3
      Top             =   210
      Width           =   5325
      Begin VB.PictureBox Picture3 
         Height          =   375
         Left            =   2160
         ScaleHeight     =   315
         ScaleWidth      =   2955
         TabIndex        =   14
         Top             =   600
         Width           =   3015
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
            Height          =   375
            Left            =   0
            TabIndex        =   15
            Top             =   0
            Width           =   3015
         End
      End
      Begin VB.PictureBox Picture2 
         Height          =   375
         Left            =   1320
         ScaleHeight     =   315
         ScaleWidth      =   555
         TabIndex        =   12
         Top             =   600
         Width           =   615
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
            Height          =   375
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Width           =   615
         End
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   375
         Left            =   120
         OleObjectBlob   =   "FRMDERAC1.frx":09C8
         TabIndex        =   11
         Top             =   600
         Width           =   1215
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
      Left            =   7800
      Picture         =   "FRMDERAC1.frx":0A4A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   885
      Visible         =   0   'False
      Width           =   855
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
      Left            =   8040
      Picture         =   "FRMDERAC1.frx":0D54
      TabIndex        =   0
      Top             =   885
      Width           =   1575
   End
   Begin VB.CommandButton command1 
      Caption         =   "Salir"
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
      Left            =   8040
      Picture         =   "FRMDERAC1.frx":105E
      TabIndex        =   1
      Top             =   270
      Width           =   1575
   End
End
Attribute VB_Name = "FRMDERAC1"
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
   USEL% = Val(Label2.Caption)
'   V& = 0
'   Call BLOQARCH("FLEXACC")
'   '
   VACON% = 1
'   FIN& = ULTREG&("FLEXACC")
'   For U& = 1 To ULTREG&("FLEXACC")
'      Call TRACE(20, U&, FIN&)
'      X$ = REGLEIDO$("FLEXACC", U&)
'      If CVI(Left$(X$, 2)) <> USEL% Then
'          V& = V& + 1
'          Call GRAREG("FLEXACC", X$, V&)
'      End If
'   Next U&
   '
   FIN& = List1.ListCount
   For U& = 1 To List1.ListCount
      Call TRACE(20, U&, FIN&)
      Z$ = List1.List(U& - 1)
      If Mid$(Z$, 45, 1) = "S" Or Mid$(Z$, 49, 1) = "S" Then
''         CAPLI$ = TRIM$(Mid$(Z$, 53, 32))
''         If Len(CAPLI$) <= 8 Then
''            Z1$ = MKI$(0) + Space$(46)
''            Call REPLA(Z1$, MKI$(USEL%), 1, 2)
''            Call REPLA$(Z1$, Mid$(Z$, 53, 8), 3, 8)
''            Call REPLA(Z1$, Mid$(Z$, 45, 1), 11, 1)
''            Call REPLA(Z1$, Mid$(Z$, 49, 1), 12, 1)
''            V& = V& + 1
''            Call GRAREG("FLEXACC", Z1$, V&)
''         End If
      End If
      '
      Call ABREADMIN
      
      CAPLI$ = TRIM$(Mid$(Z$, 53, 32))
      DRX1% = 0: If Mid$(Z$, 45, 1) = "S" Then DRX1% = 1
      DRX2% = 0: If Mid$(Z$, 49, 1) = "S" Then DRX2% = 1
      SQLX$ = "SELECT * FROM DER_ACCE "
      SQLX$ = SQLX$ + " WHERE NUM_USER = " & USEL%
      SQLX$ = SQLX$ + " AND COD_APLI = '" & CAPLI$ & "'"
      Dim PARATEMP As ADODB.Recordset
      Set PARATEMP = New ADODB.Recordset
      PARATEMP.Open SQLX$, BASSSEX, adOpenKeyset, adLockPessimistic, adCmdText
      With PARATEMP
        On Error Resume Next
        .MoveFirst
        If Err Then
           On Error GoTo 0
           If DRX1% < 1 And DRX2% < 1 Then GoTo 19
           .AddNew
         Else
           On Error GoTo 0
'           .Edit
        End If
        !NUM_USER = USEL%
        !COD_APLI = CAPLI$
        !DER_READ = DRX1%
        !DER_WRITE = DRX2%
        .Update
      End With
19    PARATEMP.Close
   Next U&
   '
   VACON% = 0
   '
   Screen.MousePointer = 1
   MsgBox "HECHO"
'   Hide
   '
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   OPPX% = COMALTER%("Seleccione el Formato del Listado:\\    Completo    \Solo Derechos Asignados")
   Call HEADERS("DERUSER", "")
   Call HEADERS("DERUSER", "Usuario: " + TRIM$(Label4) + " (" + TRIM$(Label2) + ")")
   '
   JJ& = 0
   For UI% = 1 To List1.ListCount
     XXX$ = List1.List(UI% - 1)
        If OPPX% = 1 Then
           ZZZ$ = Space$(64)
           Call REPLA(ZZZ$, Mid$(XXX$, 53, 12), 1, 12)
           Call REPLA(ZZZ$, Mid$(XXX$, 1, 36), 13, 48)
           Call REPLA(ZZZ$, Mid$(XXX$, 45, 2), 61, 2)
           Call REPLA(ZZZ$, Mid$(XXX$, 49, 2), 63, 2)
           JJ& = JJ& + 1
           Call GRAREG("DERUSER", ZZZ$, JJ&)
        ElseIf OPPX% = 2 Then
           If TRIM$(Mid$(XXX$, 45, 6)) <> "" Then
              ZZZ$ = Space$(64)
              Call REPLA(ZZZ$, Mid$(XXX$, 53, 12), 1, 12)
              Call REPLA(ZZZ$, Mid$(XXX$, 1, 36), 13, 48)
              Call REPLA(ZZZ$, Mid$(XXX$, 45, 2), 61, 2)
              Call REPLA(ZZZ$, Mid$(XXX$, 49, 2), 63, 2)
              JJ& = JJ& + 1
              Call GRAREG("DERUSER", ZZZ$, JJ&)
           End If
        Else
            Exit Sub
        End If
   Next UI%
   Call SETULTREG("DERUSER", JJ&)
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
   DESCR$ = UCase$(TRIM$(Mid$(OBX$, 17, 24)))
   '
   'VALIDO
      If TRIM$(CODI$) = "" Then
     Call COMUNI("Falta Codigo de Derecho")
     VDEF$ = OBX$: GoTo 10
   End If
   '
   If TRIM$(DESCR$) = "" Then
     Call COMUNI("Falta Descripción de Derecho")
     VDEF$ = OBX$: GoTo 10
   End If
   '
   'PASO LOS PARAMETROS A LA FUNCION DERACCE
   If DERACCE%(CODI$, DESCR$) > 0 Then
   End If
   '
   'ACTUALIZO EL LIST
   USEL% = XVALO(Label2)
   Call ACLISDER(USEL%, CODI$)
   
   Text1 = CODI$
   Text1.SetFocus
   
   
99 Command4.Enabled = True

End Sub


Private Sub Command5_Click()
   Call CARGALISTAUSUARIOS

End Sub

Private Sub Command7_Click()
    Call HELPONLINE("DERACCE")
End Sub

Private Sub Form_Load()
    Call APLICAR_SKIN(Me, App.Path & "\SKINS\dogmax.skn")
    Me.Top = mdiprincipal.Top + mdiprincipal.Height / 2 - Me.Height / 2
    Me.Left = mdiprincipal.Left + mdiprincipal.Width / 2 - Me.Width / 2
'
End Sub

Private Sub Label2_Change()
   USEL% = XVALO(Label2)
   Call ABREADMIN
   CONDI$ = "NUMEUSUARIO = " & USEL%
   DEUSER$ = VALOADMIN("USUARIOS", "NOMBRE", CONDI$, "NOMESS")
   DEUSER$ = DEUSER$ + " " + VALOADMIN("USUARIOS", "APELLIDO", CONDI$, "NOMESS")
   Label4 = DEUSER$
   Call ACLISDER1(USEL%, DEUSER$)
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
   If KeyCode = 46 Then
      CODAP$ = TRIM$(Mid$(List1.TEXT, 53, 32))
      CODAP12$ = TRIM$(Left$(CODAP$, 12))
      JJ& = 0
'      For U& = 1 To ULTREG&("INDIAPLI")
'         XX$ = REGLEIDO$("INDIAPLI", U&)
'         If TRIM$(Left$(XX$, 12)) <> CODAP12$ Then
'           JJ& = JJ& + 1
'           Call GRAREG("INDIAPLI", XX$, JJ&)
'         End If
'      Next U&
'      Call SETULTREG("INDIAPLI", JJ&)
      '
      Call ABREADMIN
      SQLX$ = "DELETE FROM INDI_APLI WHERE COD_APLI = '" & CODAP$ & "'"
      On Error Resume Next
      BaDaABRE_ADMIN_ALQ.Execute SQLX$
      If Err Then
          On Error GoTo 0
          Exit Sub
        Else
          ILISTA& = List1.ListIndex
          List1.RemoveItem ILISTA&
      End If
      On Error GoTo 0
   End If
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
Sub CARGALISTAUSUARIOS()
    Call ABREADMIN
    '
    'VALIDO POR SI NO HAY NINGUN DATO CARGADO
    CONDI$ = "NUMEUSUARIO > 0"
    CNRG& = CantRegistros&("USUARIOS", CONDI$)
    If CNRG& < 1 Then
      MsgBox "No Hay Usuarios Cargados"
      Exit Sub
    End If
    '
    JJ& = 0
    FRMBUSCAR.List1.Clear

    SQLX$ = "SELECT * FROM USUARIOS WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE NUMEUSUARIO > 0"
    SQLX$ = SQLX$ + " AND (STATUS >= 0 OR STATUS IS NULL)"
    SQLX$ = SQLX$ + " ORDER BY NUMEUSUARIO"
    '
    Set MENSATMP = BASSSEX.Execute(SQLX$)
    With MENSATMP
      Do While Not .EOF
       COCU$ = XVALO(!NUMEUSUARIO)
       DESCRI$ = SAFETEXT$(!NOMBRE) + " " + SAFETEXT$(!APELLIDO)
       Call REPLA(TTXX$, COCU$, 1, 10)
       Call REPLA(TTXX$, DESCRI$, 12, 40)
       JJ& = JJ& + 1
       FRMBUSCAR.List1.AddItem TTXX$
       .MoveNext
      Loop
    End With
    
    MENSATMP.Close
    Set MENSATMP = Nothing
    '
    FRMBUSCAR.Caption = "DERECHOS"
    FRMBUSCAR.Show

End Sub

