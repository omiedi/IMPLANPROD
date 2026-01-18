VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form ABMIMP14 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ABM DE  IMPUESTOS"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   7530
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   7335
      Begin VB.CommandButton Command2 
         Caption         =   "Rest.Imp."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         TabIndex        =   10
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton Command6 
         Caption         =   "!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   5910
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   360
         Width           =   195
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Nuevo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4440
         TabIndex        =   8
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Eliminar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Grabar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6000
         TabIndex        =   2
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   "Descripción"
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "Ir a:"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6120
         MaxLength       =   24
         TabIndex        =   0
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox TXTNOMBRE 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         Top             =   840
         Width           =   5895
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
         Height          =   225
         Index           =   0
         Left            =   0
         OleObjectBlob   =   "GESTOSUMAS.frx":0000
         TabIndex        =   11
         Top             =   0
         Visible         =   0   'False
         Width           =   1485
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   375
      Index           =   0
      Left            =   2640
      ScaleHeight     =   315
      ScaleWidth      =   1515
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   120
      OleObjectBlob   =   "GESTOSUMAS.frx":005E
      Top             =   120
   End
End
Attribute VB_Name = "ABMIMP14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Command1.Enabled = False
   '
   If XVALO(Text4) > 32767 Or XVALO(Text4) < -32767 Then
       Call COMUNI("Código Inexistente o No Válido")
       GoTo 99
   End If

   If XVALO(Text4) < 1 Then
      Call COMUNI("Identificador No Válido o Inexistente")
      GoTo 99
   End If
   '
   If TRIM$(TXTNOMBRE) = "" Then
      Call COMUNI("Imposible Grabar sin Descripción")
      GoTo 99
   End If
   
   '
   On Error GoTo ERROR_GUARDAR
   CONDI$ = "CODINT = " & CI%
   
   If CantRegistros("ITEMS", CONDI$, "NOMESS") > 0 Then
      SQLX$ = "SELECT * FROM ITEMS  WHERE CODINT = " & CI%
      Dim RS As ADODB.Recordset
      Set RS = New ADODB.Recordset
      RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic
      FLEXCONN.BeginTrans
      With RS
          !Descripcion = Left$(TXTNOMBRE, 128)
          .Update
      End With
ERROR_GUARDAR:
    If Err <> 0 Then
      FLEXCONN.RollbackTrans
      MsgBox "ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description
      GoTo 99
    Else
      FLEXCONN.CommitTrans
      Call COMUNI("Grabación Realizada Exitosamente")
      Call ACTUITEMDAT
    End If

      RS.Close
      Set RS = Nothing

   Else
      If COMALTER%("Atención !! Se va a Dar de Alta el Cliente Activo\\Continuar\Cancelar") = 1 Then
        On Error GoTo ERROR_ALTA
        Dim RS1 As ADODB.Recordset
        Set RS1 = New ADODB.Recordset
25      On Error Resume Next
        SQLX$ = "SELECT CODINT,DESCRIPCION,STATUS,MARBORRA,CODINT FROM ITEMS  WHERE CODINT =  1 "
        RS1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic
        FLEXCONN.BeginTrans
        With RS1
          .AddNew
          !CODINT = XVALO(Text4)
          !Descripcion = Left$(TXTNOMBRE, 128)
          !Status = 0
          !MARBORRA = 0
          .Update
        End With
ERROR_ALTA:
    If Err <> 0 Then
      FLEXCONN.RollbackTrans
      MsgBox "ERROR AL DAR DE ALTA,  NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description
      GoTo 99
    Else
      FLEXCONN.CommitTrans
      Call COMUNI("Alta Realizada Exitosamente")
      Call ACTUITEMDAT
    End If
        RS1.Close
        Set RS1 = Nothing
      End If

   End If
99 Command1.Enabled = True
   Text4.Locked = False
   Text4.BackColor = &HFFFFFF ' blanco
End Sub
Sub ACTUITEMDAT()
      CONDI$ = "STATUS >=0"
      Call CARGALISEL("ITEMDAT", "ITEMS", "CODINT/F5;Descripcion/A48", CONDI$, "")
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   CONDI$ = "STATUS < 0"
   Call CARGALISEL("ITEMBAJA", "ITEMS", "CODINT/F5;DESCRIPCION/A32", CONDI$, "NOMESS")
   Call SELECHO("ITEMBAJA/IMPUESTOS DATOS DE BAJA")
   '
   CI% = XVALO(VALACT1$("ITEMBAJA"))
   '
   If CI% < 1 Then GoTo 99
   On Error GoTo ERROR_GUARDAR
   CONDI$ = "CODINT = " & CI%
   
   If CantRegistros("ITEMS", CONDI$, "NOMESS") > 0 Then
      SQLX$ = "SELECT * FROM ITEMS  WHERE CODINT = " & CI%
      Dim RS As ADODB.Recordset
      Set RS = New ADODB.Recordset
      RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic
      FLEXCONN.BeginTrans
      With RS
          !Status = 0
          .Update
      End With
      
ERROR_GUARDAR:
      If Err <> 0 Then
          FLEXCONN.RollbackTrans
          MsgBox "ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description
          
      Else
          FLEXCONN.CommitTrans
          Call COMUNI("Alta Realizada Exitosamente")
          Call ACTUITEMDAT
      End If
    
          RS.Close
          Set RS = Nothing
    Else
      Call COMUNI("Código Inexistente o no Válido")
    End If
    
99  Command2.Enabled = True

End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   CI% = XVALO(Text4)
   '
   If XVALO(CI%) < 1 Then
      Call COMUNI("Identificador No Válido o Inexistente")
      GoTo 99
   End If
   '
   On Error GoTo ERROR_GUARDAR
   CONDI$ = "CODINT = " & CI%
   
   If CantRegistros("ITEMS", CONDI$, "NOMESS") > 0 Then
      SQLX$ = "SELECT * FROM ITEMS  WHERE CODINT = " & CI%
      Dim RS As ADODB.Recordset
      Set RS = New ADODB.Recordset
      RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic
      FLEXCONN.BeginTrans
      With RS
          !Status = -1
          .Update
      End With
      
ERROR_GUARDAR:
      If Err <> 0 Then
          FLEXCONN.RollbackTrans
          MsgBox "ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description
          
      Else
          FLEXCONN.CommitTrans
          Call COMUNI("Baja Realizada Exitosamente")
          Call ACTUITEMDAT
      End If
    
          RS.Close
          Set RS = Nothing
    Else
      Call COMUNI("Código Inexistente o no Válido")
    End If
    
99  Command4.Enabled = True
End Sub




Private Sub Command6_Click()
    Command6.Enabled = False
    Call SELECHO("ITEMDAT/IMPUESTOS")
    CI1% = XVALO(VALACT1$("ITEMDAT"))
    Text4 = CI1%
    Unload FORSELEC
    Command6.Enabled = True

End Sub

Private Sub Command9_Click()
   Command9.Enabled = False
   Call BLANFORMU
   If XVALO(Text4) > 32767 Or XVALO(Text4) < -32767 Then
       Call COMUNI("Código Inexistente o No Válido")
       GoTo 99
   End If

   CONDI$ = "CODINT > 0"
   CI% = 1 + XVALO(VALOBADA("ITEMS", "MAX(CODINT)", CONDI$, "NOMESS"))
   '
   Text4 = CI%
   Text4.Locked = True
   Text4.BackColor = &H8000000F
   '
99 Command9.Enabled = True
End Sub

Private Sub Form_Load()

'   Call APLICAR_SKIN(Me, LUDAT$ & "\SKINS\dogmax.skn")
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)

    '
End Sub

Private Sub Text4_DblClick()
   Text4.Locked = False
   Text4.BackColor = &HFFFFFF ' blanco

End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If
End Sub




Private Sub Text4_Change()
   Call BLANFORMU
   
   If XVALO(Text4) > 32767 Or XVALO(Text4) < -32767 Then
       Call COMUNI("Código Inexistente o No Válido")
       Exit Sub
   End If
   
   CI% = XVALO(Text4)
   If CI% < 1 Then Exit Sub
   '
   CONDI$ = "CODINT = " & CI% & " AND STATUS >= 0 "
   If CantRegistros("ITEMS", CONDI$, "NOMESS") > 0 Then
      SQLX$ = "SELECT * FROM ITEMS WITH(NOLOCK) WHERE CODINT = " & CI%
      Set RS = READSET(SQLX$)
      With RS
        TXTNOMBRE = SAFETEXT$(!Descripcion)
      End With

   End If
End Sub

Sub BLANFORMU()
    TXTNOMBRE = ""
End Sub


Private Sub TXTNOMBRE_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If

End Sub

