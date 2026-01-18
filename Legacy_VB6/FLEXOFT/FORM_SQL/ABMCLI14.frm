VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form ABMCLI14 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ABM DE CLIENTES "
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   10170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   10170
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   36
      TabStop         =   0   'False
      Text            =   "Alta - Baja - Modificación de Clientes"
      Top             =   120
      Width           =   9960
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Cliente"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   120
      TabIndex        =   14
      Top             =   480
      Width           =   9975
      Begin VB.TextBox TXTIVA 
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
         TabIndex        =   10
         Top             =   2280
         Width           =   615
      End
      Begin VB.CommandButton Command3 
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
         Left            =   6720
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   2280
         Width           =   195
      End
      Begin VB.TextBox Text2 
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
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   2280
         Width           =   2655
      End
      Begin VB.TextBox TEXT17 
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
         Index           =   4
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         Text            =   "C.U.I.T."
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Rest.Clientes"
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
         TabIndex        =   31
         Top             =   3240
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
         Left            =   3630
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   360
         Width           =   195
      End
      Begin VB.TextBox TXTTELEFONO1 
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
         TabIndex        =   4
         Top             =   2280
         Width           =   3495
      End
      Begin VB.TextBox TEXT17 
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
         Index           =   3
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Text            =   "Iva "
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox TXTCUIT 
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
         TabIndex        =   11
         Top             =   2760
         Width           =   3495
      End
      Begin VB.TextBox TEXT17 
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
         Index           =   2
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Text            =   "Contacto"
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox TXTCONTACTO 
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
         TabIndex        =   9
         Top             =   1800
         Width           =   3495
      End
      Begin VB.TextBox TEXT17 
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
         Index           =   1
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Text            =   "Cód.Postal"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox TXTCODIGOPOSTAL 
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
         TabIndex        =   8
         Top             =   1320
         Width           =   3495
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
         Left            =   7080
         TabIndex        =   26
         Top             =   3240
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
         TabIndex        =   25
         Top             =   3240
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
         Left            =   8400
         TabIndex        =   12
         Top             =   3240
         Width           =   1215
      End
      Begin VB.TextBox TEXT17 
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
         Index           =   0
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Text            =   "Email 2"
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox TXTEMAIL2 
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
         TabIndex        =   7
         Top             =   840
         Width           =   3495
      End
      Begin VB.TextBox Text15 
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
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Text            =   "Email 1"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox TXTEMAIL1 
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
         TabIndex        =   6
         Top             =   360
         Width           =   3495
      End
      Begin VB.TextBox Text13 
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
         TabIndex        =   22
         TabStop         =   0   'False
         Text            =   "Teléfono 2 "
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox TXTTELEFONO2 
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
         TabIndex        =   5
         Top             =   2760
         Width           =   3495
      End
      Begin VB.TextBox Text12 
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
         TabIndex        =   21
         TabStop         =   0   'False
         Text            =   "Teléfono 1 "
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox Text10 
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
         TabIndex        =   20
         TabStop         =   0   'False
         Text            =   "Localidad"
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox TXTLOCALIDAD 
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
         TabIndex        =   3
         Top             =   1800
         Width           =   3495
      End
      Begin VB.TextBox Text8 
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
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   "Domicilio"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox TXTDOMICILIO 
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
         TabIndex        =   2
         Top             =   1320
         Width           =   3495
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
         TabIndex        =   18
         TabStop         =   0   'False
         Text            =   "Nombre"
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
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   17
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
         Left            =   3840
         MaxLength       =   24
         TabIndex        =   0
         Top             =   360
         Width           =   975
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
         Width           =   3495
      End
      Begin VB.TextBox Text1 
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
         TabIndex        =   16
         TabStop         =   0   'False
         Text            =   "ID"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox TXTID 
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
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   24
         TabIndex        =   15
         Top             =   360
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
         Height          =   225
         Index           =   0
         Left            =   0
         OleObjectBlob   =   "ABMCLI14.frx":0000
         TabIndex        =   32
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
      TabIndex        =   13
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   120
      OleObjectBlob   =   "ABMCLI14.frx":005E
      Top             =   120
   End
End
Attribute VB_Name = "ABMCLI14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Command1.Enabled = False
   '
   If XVALO(TXTID) > 32767 Or XVALO(TXTID) < -32767 Then
      Call COMUNI("Identificador No Válido o Inexistente")
      GoTo 99
   End If
   '
   NC% = XVALO(TXTID)
   '
   If XVALO(TXTID) < 1 Then
      Call COMUNI("Identificador No Válido o Inexistente")
      GoTo 99
   End If
   
   If TRIM$(TXTNOMBRE) = "" Then
     Call COMUNI("Imposible Guardar Cliente sin Nombre")
     GoTo 99
   End If
   '
   On Error GoTo ERROR_GUARDAR
   CONDI$ = "Nume_Cli = " & NC%
   
   If CantRegistros("DEUDOR12", CONDI$, "NOMESS") > 0 Then
      SQLX$ = "SELECT * FROM DEUDOR12  WHERE NUME_CLI = " & NC%
      Dim RS As ADODB.Recordset
      Set RS = New ADODB.Recordset
      RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic
      FLEXCONN.BeginTrans
      With RS
          !Raso_Cli = Left$(TXTNOMBRE, 48)
          !Domi_Cli = Left$(TXTDOMICILIO, 64)
          !Loca_Cli = Left$(TXTLOCALIDAD, 48)
          !TELE_Cli = Left$(TXTTELEFONO1, 48)
          !FAXI_Cli = Left$(TXTTELEFONO2, 32)
          !Mail_Cli = Left$(TXTEMAIL1, 64)
          !PagWeb = Left$(TXTEMAIL2, 128)
          !Cpos_Cli = Left$(TXTCODIGOPOSTAL, 16)
          !Ctac_Cli = Left$(TXTCONTACTO, 48)
          !Cuit_Cli = Left$(TXTCUIT, 24)
          !Piva_Cli = XVALO(TXTIVA)
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
    End If

      RS.Close
      Set RS = Nothing

   Else
      If COMALTER%("Atención !! Se va a Dar de Alta el Cliente Activo\\Continuar\Cancelar") = 1 Then
        On Error GoTo ERROR_ALTA
        Dim RS1 As ADODB.Recordset
        Set RS1 = New ADODB.Recordset
25      On Error Resume Next
        SQLX$ = "SELECT * FROM DEUDOR12  WHERE NUME_CLI =  1 "
        RS1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic
        FLEXCONN.BeginTrans
        With RS1
          .AddNew
          !NUME_CLI = NC%
          !Raso_Cli = Left$(TXTNOMBRE, 48)
          !Domi_Cli = Left$(TXTDOMICILIO, 64)
          !Loca_Cli = Left$(TXTLOCALIDAD, 48)
          !TELE_Cli = Left$(TXTTELEFONO1, 48)
          !FAXI_Cli = Left$(TXTTELEFONO2, 32)
          !Mail_Cli = Left$(TXTEMAIL1, 64)
          !PagWeb = Left$(TXTEMAIL2, 128)
          !Cpos_Cli = Left$(TXTCODIGOPOSTAL, 16)
          !Ctac_Cli = Left$(TXTCONTACTO, 48)
          !Cuit_Cli = Left$(TXTCUIT, 24)
          !Piva_Cli = XVALO(TXTIVA)
          !STAT_CLI = 1
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
    End If
        RS1.Close
        Set RS1 = Nothing
      End If
    Call GENECOCLI
    Text4.Locked = False
    Text4.BackColor = &HFFFFFF ' blanco

   End If
99 Command1.Enabled = True
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   Call CARGALISEL("CLIEBAJA", "DEUDOR12", "NUME_CLI/f5;Raso_cli/A32", "STAT_CLI < 0", "NOMESS")
   Call SELECHO("CLIEBAJA/CLIENTES DE BAJA")
   NC% = XVALO(VALACT1$("CLIEBAJA/CLIENTES DE BAJA"))
   If NC% = 0 Then GoTo 99
   '
   On Error GoTo ERROR_GUARDAR
   CONDI$ = "Nume_Cli = " & NC%
   
   If CantRegistros("DEUDOR12", CONDI$, "NOMESS") > 0 Then
      SQLX$ = "SELECT * FROM DEUDOR12  WHERE NUME_CLI = " & NC%
      Dim RS As ADODB.Recordset
      Set RS = New ADODB.Recordset
      RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic
      FLEXCONN.BeginTrans
      With RS
          !STAT_CLI = 0
          .Update
      End With
      
ERROR_GUARDAR:
      If Err <> 0 Then
          FLEXCONN.RollbackTrans
          MsgBox "ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description
          GoTo 99
      Else
          FLEXCONN.CommitTrans
          Call COMUNI("Restauración Realizada Exitosamente")
          Text4 = ""
          Text4 = NC%
      End If
    
          RS.Close
          Set RS = Nothing
    Else
      Call COMUNI("Código Inexistente o no Válido")
    End If
    Call GENECOCLI
    Text4.Locked = False
    Text4.BackColor = &HFFFFFF ' blanco

99 Command2.Enabled = True

End Sub

Private Sub Command3_Click()
   Call SELECHO("CAPOSIV")
   VDEVU$ = VALACT1$("CAPOSIV")
   If TRIM$(VDEVU$) <> "" Then
      TXTIVA = VDEVU$
   End If

End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   NC% = XVALO(TXTID)
   '
   If XVALO(TXTID) < 1 Then
      Call COMUNI("Identificador No Válido o Inexistente")
      GoTo 99
   End If
   '
   On Error GoTo ERROR_GUARDAR
   CONDI$ = "Nume_Cli = " & NC%
   
   If CantRegistros("DEUDOR12", CONDI$, "NOMESS") > 0 Then
      SQLX$ = "SELECT * FROM DEUDOR12  WHERE NUME_CLI = " & NC%
      Dim RS As ADODB.Recordset
      Set RS = New ADODB.Recordset
      RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic
      FLEXCONN.BeginTrans
      With RS
          !STAT_CLI = -1
          .Update
      End With
      
ERROR_GUARDAR:
      If Err <> 0 Then
          FLEXCONN.RollbackTrans
          MsgBox "ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description
          GoTo 99
      Else
          FLEXCONN.CommitTrans
          Call COMUNI("Baja Realizada Exitosamente")
          Call BLANFORMU
      End If
    
          RS.Close
          Set RS = Nothing
    Else
      Call COMUNI("Código Inexistente o no Válido")
    End If
    '
    Call GENECOCLI
    Text4.Locked = False
    Text4.BackColor = &HFFFFFF ' blanco


99  Command4.Enabled = True
End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
    Call SELECHO("DEUDOR1")
    NCLIE% = XVALO(TRIM$(VALACT1$("DEUDOR1")))
    If NCLIE% > 0 Then
      Text4.TEXT = TRIM$(Str$(NCLIE%))
    End If
    Command6.Enabled = True

End Sub
Sub GENECOCLI(Optional LONGINT%)
      Screen.MousePointer = 11
      Call CARGALISEL("DEUDOR1", "Deudor12", "Nume_Cli/F6.0>;Raso_Cli/A48", "Stat_Cli >= 0 and Nume_Cli>0 ORDER BY Raso_Cli")
      Screen.MousePointer = 1
End Sub

Private Sub Command9_Click()
   Command9.Enabled = False
   Call BLANFORMU
   CONDI$ = "Nume_Cli >0"
   NC% = 1 + XVALO(VALOBADA("DEUDOR12", "MAX(Nume_Cli)", CONDI$, "NOMESS"))
   '
   Text4 = NC%
   '
   Text4.Locked = True
   Text4.BackColor = &H8000000F

   Command9.Enabled = True
End Sub

Private Sub Form_Load()
'Call APLICAR_SKIN(Me, LUDAT$ & "\SKINS\dogmax.skn")
UTILIZA_SKIN% = 1
Call APLICACIONSKINS(Me)
    '
End Sub
Private Sub mnuListcli_Click()
     Call FINAL("?COBCONTAR") ' cons
End Sub

Private Sub TXTIVA_Change()
   Text2 = ECOARCH$("CAPOSIV", Chr(XVALO(TXTIVA)))
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If
End Sub

Private Sub TXTCODIGOPOSTAL_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If

End Sub

Private Sub TXTCONTACTO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If

End Sub

Private Sub TXTDOMICILIO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If

End Sub

Private Sub TXTEMAIL1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If

End Sub

Private Sub TXTEMAIL2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If

End Sub

Private Sub TXTID_Change()
   Call BLANFORMU
   If XVALO(TXTID) > 32767 Then Exit Sub
   NC% = XVALO(TXTID)
   If NC% < 1 Then Exit Sub
   '
   CONDI$ = "Nume_Cli = " & NC% & " And Stat_Cli >= 0"
   If CantRegistros("DEUDOR12", CONDI$, "NOMESS") > 0 Then
   
      SQLX$ = "SELECT * FROM DEUDOR12 WITH(NOLOCK) WHERE NUME_CLI = " & NC% & " And Stat_Cli >= 0"
      Set RS = READSET(SQLX$)
      With RS
        If Not .EOF Then
            TXTNOMBRE = SAFETEXT$(!Raso_Cli)
            TXTDOMICILIO = SAFETEXT$(!Domi_Cli)  '64
            TXTLOCALIDAD = SAFETEXT$(!Loca_Cli)  '48
            TXTTELEFONO1 = SAFETEXT$(!TELE_Cli)  '48
            TXTTELEFONO2 = SAFETEXT$(!FAXI_Cli)  '32
            TXTEMAIL1 = SAFETEXT$(!Mail_Cli)     '64
            TXTEMAIL2 = SAFETEXT$(!PagWeb)       '128
            TXTCODIGOPOSTAL = SAFETEXT$(!Cpos_Cli) '16
            TXTCONTACTO = SAFETEXT$(!Ctac_Cli)     '48
            TXTCUIT = SAFETEXT$(!Cuit_Cli)    '24
            TXTIVA = XVALO(!Piva_Cli)
        End If
      End With
   End If
End Sub

Private Sub Text4_Change()
   TXTID = Text4
   
End Sub
Sub BLANFORMU()
        TXTNOMBRE = ""
        TXTDOMICILIO = "" '64
        TXTLOCALIDAD = "" '48
        TXTTELEFONO1 = ""  '48
        TXTTELEFONO2 = ""  '32
        TXTEMAIL1 = ""     '64
        TXTEMAIL2 = ""     '128
        TXTCODIGOPOSTAL = "" '16
        TXTCONTACTO = ""    '48
        TXTCUIT = ""   '24
        TXTIVA = ""
End Sub

Private Sub TXTIVA_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If

End Sub

Private Sub TXTLOCALIDAD_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If

End Sub

Private Sub TXTNOMBRE_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If

End Sub

Private Sub TXTTELEFONO1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If

End Sub

Private Sub TXTTELEFONO2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If

End Sub
Private Sub Text4_DblClick()
   Text4.Locked = False
   Text4.BackColor = &HFFFFFF ' blanco

End Sub

