VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form ABMUSUARIOS 
   Caption         =   "ABM USUARIOS"
   ClientHeight    =   5310
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7035
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   7035
   Begin VB.CheckBox Check1 
      Caption         =   "                    Otorgar Derechos de Supervisor"
      Height          =   735
      Left            =   2040
      TabIndex        =   17
      Top             =   4440
      Width           =   2055
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   1440
      OleObjectBlob   =   "ABMUSUARIOS.frx":0000
      Top             =   4560
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Guardar "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   6
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Nuevo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5760
      TabIndex        =   8
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "ABM USUARIOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6855
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1800
         MaxLength       =   12
         TabIndex        =   2
         Top             =   2160
         Width           =   4935
      End
      Begin VB.CommandButton Command1 
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
         Left            =   4560
         TabIndex        =   7
         Top             =   480
         Width           =   255
      End
      Begin VB.TextBox Text8 
         Height          =   885
         Left            =   1800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   3120
         Width           =   4935
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1800
         MaxLength       =   12
         TabIndex        =   3
         Top             =   2640
         Width           =   4935
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1800
         MaxLength       =   64
         TabIndex        =   1
         Top             =   1680
         Width           =   4935
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1800
         MaxLength       =   64
         TabIndex        =   0
         Top             =   1200
         Width           =   4935
      End
      Begin VB.PictureBox Picture1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4800
         ScaleHeight     =   315
         ScaleWidth      =   1875
         TabIndex        =   10
         Top             =   480
         Width           =   1935
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
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
            Height          =   375
            Left            =   0
            TabIndex        =   11
            Top             =   0
            Width           =   1935
         End
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   3720
         OleObjectBlob   =   "ABMUSUARIOS.frx":0234
         TabIndex        =   12
         Top             =   480
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   400
         OleObjectBlob   =   "ABMUSUARIOS.frx":02A0
         TabIndex        =   13
         Top             =   1200
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   400
         OleObjectBlob   =   "ABMUSUARIOS.frx":030C
         TabIndex        =   14
         Top             =   1680
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   405
         OleObjectBlob   =   "ABMUSUARIOS.frx":037C
         TabIndex        =   15
         Top             =   2640
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   405
         OleObjectBlob   =   "ABMUSUARIOS.frx":03E6
         TabIndex        =   16
         Top             =   3120
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   400
         OleObjectBlob   =   "ABMUSUARIOS.frx":0460
         TabIndex        =   18
         Top             =   2160
         Width           =   1335
      End
   End
End
Attribute VB_Name = "ABMUSUARIOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
Private Sub Command1_Click()
    Command1.Enabled = False
    
    Call ABREADMIN
    '
    'VALIDO POR SI NO HAY NINGUN DATO CARGADO
    CONDI$ = "NUMEUSUARIO > 0"
    CNRG& = CantRegistros&("USUARIOS", CONDI$)
    If CNRG& < 1 Then
      MsgBox "No Hay Usuarios Cargados"
      GoTo 99
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
    'Set MENSATMP = BaDaABRE_ADMIN_ALQ.OpenRecordset(SQLX$, dbOpenSnapshot)
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
    FRMBUSCAR.Caption = "USUARIOS"
    FRMBUSCAR.Show 1
99  Command1.Enabled = True

End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   NUEVOUSUARIO.Show 1
   
   Command2.Enabled = True

End Sub

Private Sub Command3_Click()

    Command3.Enabled = False
    Call ABREADMIN
    JJ& = 0
    '
    'TOMO EL NUMERO DEL LABEL
    CUENTA& = XVALO(Label9)
    '
    'VALIDACION QUE SE HAYA SELECCIONADO UN CODIGO
    If CUENTA& < 1 Then
       MsgBox "Código de Cuenta No Valido"
       GoTo 99
    End If
    '
    If TRIM(Text2) = "" Then
       MsgBox "Debe Ingresar Nombre"
       GoTo 99
    End If
    '
    If TRIM(Text1) = "" Then
       MsgBox "Debe Ingresar Usuario"
       GoTo 99
    End If

    'VALIDAR SI EL USUARIO ES EXISTENTE
    CONDI$ = "USUARIO =  '" & UCase$(TRIM$(Text1)) & "' AND NUMEUSUARIO <> " & XVALO(Label9)
    If CantRegistros&("USUARIOS", CONDI$) > 0 Then
       MsgBox "USUARIO EXISTENTE " + vbCrLf + "FAVOR DE CAMBIAR Y VOLVER A GRABAR"
       GoTo 99
    End If
    
    'GRABACION , SI NO EXISTE LO GENERA SI NO LO EDITA
    SQLX$ = "SELECT * FROM USUARIOS "
    SQLX$ = SQLX$ + " WHERE NUMEUSUARIO = " & CUENTA&
    '
'''    Set RECTMP = BaDaABRE_ADMIN_ALQ.OpenRecordset(SQLX$, dbOpenDynaset)
'''
'''    With RECTMP
'''          If (.BOF And .EOF) Then
'''            .AddNew
'''            !NUMEUSUARIO = CUENTA&
'''          Else
'''            .Edit
'''          End If
'''            !NOMBRE = TRIM$(Text2)
'''            !APELLIDO = TRIM$(Text3)
'''            !CLAVE = TRIM$(Text4)
'''            !USUARIO = UCase$(TRIM$(Text1))
'''            !OBSERVACIONES = Text8
'''            !DERE_SUPERVISOR = Check1.Value
'''            .Update
'''          End With
    '

    On Error GoTo ERROR_GUARDAR
    Set AUX = New ADODB.Recordset
    AUX.Open (SQLX$), BASSSEX, adOpenKeyset, adLockPessimistic, adCmdText
    BASSSEX.BeginTrans
    With AUX
'          If (.BOF And .EOF) Then COMENTADO POR QUE SE CAMBIO LA FOMRA DE DAR DE ALTA
'            .AddNew
'            !NUMEUSUARIO = CUENTA&
'          End If
            !NOMBRE = TRIM$(Text2)
            !APELLIDO = TRIM$(Text3)
            !CLAVE = TRIM$(Text4)
            !USUARIO = UCase$(TRIM$(Text1))
            !OBSERVACIONES = Text8
            !DERE_SUPERVISOR = Check1.Value
            .Update
    End With
    
ERROR_GUARDAR:
    If Err <> 0 Then
      BASSSEX.RollbackTrans
      MsgBox "ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description
    Else
      BASSSEX.CommitTrans
      AUX.Close
      Set AUX = Nothing
      MsgBox "HECHO"
    End If
    
    
99  Command3.Enabled = True
End Sub

Private Sub Command4_Click()

    Command4.Enabled = False
    Call ABREADMIN
    JJ& = 0
    '
    'TOMO EL NUMERO DEL LABEL
    CUENTA& = XVALO(Label9)
    '
    'VALIDACION QUE SE HAYA SELECCIONADO UN CODIGO
    If CUENTA& < 1 Then
       MsgBox "Código de Usuario No Valido"
       GoTo 99
    End If
    '
    If CUENTA& = 1 Then
       MsgBox "IMPOSIBLE ELIMINAR USUARIO 1 CONSIDERADO SUPERVISOR DEL SISTEMA": GoTo 99
    End If
    CONDI$ = "NUMEUSUARIO = " & CUENTA&
    '
    'VALIDO SI EXISTE
    CNRG& = CantRegistros&("USUARIOS", CONDI$)
    If CNRG& < 1 Then
      MsgBox "Código de Usuario No Valido"
      GoTo 99
    End If
    '
    TTXX$ = "Realmente desea dar de Baja el Usuario Activo"
    If MsgBox(TTXX$, vbYesNo, "Baja de Usuario") <> vbYes Then GoTo 99
    'GRABACION , SI NO EXISTE LO GENERA SI NO LO EDITA
    SQLX$ = "SELECT * FROM USUARIOS "
    SQLX$ = SQLX$ + " WHERE NUMEUSUARIO = " & CUENTA&
    '
    Dim RECTMP As ADODB.Recordset
    Set RECTMP = New ADODB.Recordset
    RECTMP.Open SQLX$, BASSSEX, adOpenKeyset, adLockPessimistic, adCmdText
    'Set RECTMP = BaDaABRE_ADMIN_ALQ.OpenRecordset(SQLX$, dbOpenDynaset)
    With RECTMP
        '.Edit
        !Status = -1
        .Update
    End With
    '
    RECTMP.Close
    Set RECTMP = Nothing
    '
    Call LIMPIAR
    Label9 = ""
    
99  Command4.Enabled = True

End Sub

Private Sub Form_Load()

    ABMUSUARIOS.Height = 5820
    ABMUSUARIOS.Top = 1000
    ABMUSUARIOS.Width = 7155
    ABMUSUARIOS.Left = 1000
    Call APLICAR_SKIN(Me, App.Path & "\SKINS\dogmax.skn")
    '
        
    Me.Top = mdiprincipal.Top + mdiprincipal.Height / 2 - Me.Height / 2
    Me.Left = mdiprincipal.Left + mdiprincipal.Width / 2 - Me.Width / 2

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Unload Me
End Sub

Private Sub Label9_Change()

   CUENTA& = XVALO(Label9)
   If CUENTA& < 1 Then Exit Sub
   '
   CONDI$ = "NUMEUSUARIO = " & CUENTA&
   '
   'VALIDO SI EXISTE
   CNRG& = CantRegistros&("USUARIOS", CONDI$)
   If CNRG& < 1 Then Call LIMPIAR: Exit Sub
   
   '
   Text2 = VALOADMIN("USUARIOS", "NOMBRE", CONDI$, "NOMESS")
   Text3 = VALOADMIN("USUARIOS", "APELLIDO", CONDI$, "NOMESS")
   Text4 = VALOADMIN("USUARIOS", "CLAVE", CONDI$, "NOMESS")
   Text8 = VALOADMIN("USUARIOS", "OBSERVACIONES", CONDI$, "NOMESS")
   Text1 = VALOADMIN("USUARIOS", "USUARIO", CONDI$, "NOMESS")
   Check1.Value = XVALO(VALOADMIN("USUARIOS", "DERE_SUPERVISOR", CONDI$, "NOMESS"))
   '
End Sub


Sub LIMPIAR()
   Text1 = ""
   Text2 = ""
   Text3 = ""
   Text4 = ""
   Text8 = ""
   Check1.Value = 0
   
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

