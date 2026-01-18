VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form NUEVOUSUARIO1 
   Caption         =   "Nuevo"
   ClientHeight    =   2730
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   ScaleHeight     =   2730
   ScaleWidth      =   6165
   StartUpPosition =   3  'Windows Default
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   1440
      OleObjectBlob   =   "NUEVOUSUARIO1.frx":0000
      Top             =   2160
   End
   Begin VB.Frame Frame1 
      Caption         =   "Nuevo Usuario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6135
      Begin VB.CommandButton Command3 
         Caption         =   "DAR DE ALTA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   4
         Top             =   2160
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1080
         MaxLength       =   64
         TabIndex        =   0
         Top             =   360
         Width           =   4935
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1080
         MaxLength       =   64
         TabIndex        =   1
         Top             =   840
         Width           =   4935
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1080
         MaxLength       =   12
         TabIndex        =   3
         Top             =   1800
         Width           =   4935
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1080
         MaxLength       =   12
         TabIndex        =   2
         Top             =   1320
         Width           =   4935
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   -315
         OleObjectBlob   =   "NUEVOUSUARIO1.frx":0234
         TabIndex        =   6
         Top             =   360
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   -315
         OleObjectBlob   =   "NUEVOUSUARIO1.frx":02A0
         TabIndex        =   7
         Top             =   840
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   -315
         OleObjectBlob   =   "NUEVOUSUARIO1.frx":0310
         TabIndex        =   8
         Top             =   1800
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   -315
         OleObjectBlob   =   "NUEVOUSUARIO1.frx":037A
         TabIndex        =   9
         Top             =   1320
         Width           =   1335
      End
   End
End
Attribute VB_Name = "NUEVOUSUARIO1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()
    Command3.Enabled = False
    '
    Call ABRECONEXION
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
    '
    NUMPRO& = 1 + MAXNUMCAMPO&("USUARIOS", "NUMEUSUARIO")

    'GRABACION , SI NO EXISTE LO GENERA SI NO LO EDITA
    SQLX$ = "SELECT * FROM USUARIOS "
    SQLX$ = SQLX$ + " WHERE NUMEUSUARIO = " & NUMPRO&
    
    On Error GoTo ERROR_GUARDAR
    Set AUX = New ADODB.Recordset
    AUX.Open (SQLX$), BASSSEX, adOpenKeyset, adLockPessimistic, adCmdText
    BASSSEX.BeginTrans
    With AUX
            .AddNew
            !NUMEUSUARIO = NUMPRO&
            !nombre = TRIM$(Text2)
            !APELLIDO = TRIM$(Text3)
            !CLAVE = TRIM$(Text4)
            !USUARIO = UCase$(TRIM$(Text1))
            .Update
    End With
    
ERROR_GUARDAR:
    If Err <> 0 Then
      BASSSEX.RollbackTrans
      MsgBox "ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description
      GoTo 99
    Else
      BASSSEX.CommitTrans
      'LO PASO ASI POR QUE SI SOLO PASO EL NUMERO AL LABEL9 NO EJECUTA EL EVENTO CHANGE
      ABMUSUARIOS.Label9 = Str(NUMPRO&)
      ABMUSUARIOS.Text2 = TRIM$(Text2)
      ABMUSUARIOS.Text3 = TRIM$(Text3)
      ABMUSUARIOS.Text4 = TRIM$(Text4)
      ABMUSUARIOS.Text1 = UCase$(TRIM$(Text1))
      Unload Me
    End If
    
    
    AUX.Close
    Set AUX = Nothing
    MsgBox "HECHO"
99  Command3.Enabled = True

End Sub

Private Sub Form_Load()
    Call APLICAR_SKIN(Me, App.Path & "\SKINS\dogmax.skn")
    Me.Top = mdiprincipal.Top + mdiprincipal.Height / 2 - Me.Height / 2
    Me.Left = mdiprincipal.Left + mdiprincipal.Width / 2 - Me.Width / 2
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

Private Sub Text4_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If
End Sub
