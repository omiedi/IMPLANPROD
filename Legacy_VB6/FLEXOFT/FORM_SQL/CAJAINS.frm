VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Begin VB.Form CAJAINS 
   Caption         =   "Movimientos de Caja"
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   14610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   14610
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtingreso 
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
      Left            =   7320
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   7605
      Width           =   975
   End
   Begin VB.TextBox Text16 
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
      Index           =   5
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   7605
      Width           =   3615
   End
   Begin VB.TextBox txtgastos 
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
      Left            =   7320
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   6840
      Width           =   975
   End
   Begin VB.TextBox Text16 
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
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   6840
      Width           =   3615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Actualizar Información"
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
      Left            =   120
      TabIndex        =   18
      Top             =   840
      Width           =   3255
   End
   Begin VB.PictureBox CommonDialog1 
      Height          =   480
      Left            =   6120
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   21
      Top             =   3960
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.PictureBox Skin1 
      Height          =   480
      Left            =   8520
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   22
      Top             =   720
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton Command25 
      Caption         =   "Imprimir"
      Height          =   495
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7800
      Width           =   2415
   End
   Begin VB.TextBox Text20 
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
      Index           =   1
      Left            =   2760
      MaxLength       =   50
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   7605
      Width           =   4455
   End
   Begin VB.TextBox Text20 
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
      Index           =   0
      Left            =   2760
      MaxLength       =   50
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   6840
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Grabar Retiro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      TabIndex        =   16
      Top             =   7200
      Width           =   2175
   End
   Begin VB.TextBox Text2 
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
      Index           =   1
      Left            =   1680
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   7605
      Width           =   975
   End
   Begin VB.TextBox Text16 
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
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Text            =   "Retiro de Caja:"
      Top             =   6840
      Width           =   1455
   End
   Begin VB.TextBox Text2 
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
      Index           =   0
      Left            =   1680
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   6840
      Width           =   975
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
      Index           =   3
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Text            =   "Ingreso de Caja:"
      Top             =   7605
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Grabar Ingreso Manual"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      TabIndex        =   11
      Top             =   8040
      Width           =   2175
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
      Index           =   8
      Left            =   12120
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "Salidas:"
      Top             =   6900
      Width           =   1095
   End
   Begin VB.TextBox Text16 
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
      Left            =   12120
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "Resultado:"
      Top             =   7320
      Width           =   1095
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
      Index           =   0
      Left            =   13440
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   6480
      Width           =   1095
   End
   Begin VB.TextBox Text16 
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
      Left            =   12120
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "Ingreso:"
      Top             =   6480
      Width           =   1095
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
      Index           =   1
      Left            =   13440
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   6900
      Width           =   1095
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
      Index           =   2
      Left            =   13440
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   7320
      Width           =   1095
   End
   Begin VB.TextBox Text17 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FBCA9D&
      Height          =   600
      Index           =   3
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "MOVIMIENTOS DE CAJA"
      Top             =   0
      Width           =   14595
   End
   Begin VB.Frame Frame1 
      Height          =   1320
      Left            =   0
      TabIndex        =   1
      Top             =   620
      Width           =   14535
      Begin VB.TextBox TXTDESCRIUSUARIO 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   720
         Width           =   4815
      End
      Begin VB.TextBox TXTNUMUSUARIO 
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
         Height          =   375
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Selección de Usuarios"
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
         Left            =   120
         TabIndex        =   26
         Top             =   720
         Width           =   3255
      End
      Begin VB.TextBox Text16 
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
         Height          =   375
         Index           =   3
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Text            =   "Registros Generados"
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox TXTCANTIFILAS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker Desde 
         Height          =   330
         Left            =   11280
         TabIndex        =   24
         Top             =   240
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   128057345
         CurrentDate     =   41123
      End
      Begin MSComCtl2.DTPicker Hasta 
         Height          =   330
         Left            =   12960
         TabIndex        =   25
         Top             =   240
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   128057345
         CurrentDate     =   41123
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   4455
      Left            =   0
      TabIndex        =   23
      Top             =   1920
      Width           =   14520
      _ExtentX        =   25612
      _ExtentY        =   7858
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Menu MnuArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mnucajadetalle 
         Caption         =   "Detalle de Caja"
      End
      Begin VB.Menu mnuTRansfxUsuario 
         Caption         =   "Transferencias de caja entre usuarios"
      End
   End
End
Attribute VB_Name = "CAJAINS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click(Index As Integer)
   Command1(Index).Enabled = False
   USNX% = USNBR% Mod 100
   If USNX% <> 1 Then
      Call COMUNI("Solo el Usuario Supervisor Puede Realizar Retiros o Ingresos a Caja")
      GoTo 99
   End If
   If TRIM$(Text20(Index)) = "" Then
        MsgBox "FALTA INGRESAR REFERENCIA"
        GoTo 99
   End If

   Select Case (Index)
     Case 0
        If XVALO(Text2(Index)) > 0 Then Call GRABA_MOVI_CAJA(XVALO(Text2(0)), 0)
     Case 1
        If XVALO(Text2(Index)) > 0 Then Call GRABA_MOVI_CAJA(0, XVALO(Text2(1)))
   End Select
'   MsgBox "HECHO"
   Call Command2_Click
   Text2(0) = ""
   Text2(1) = ""
   Text20(0) = ""
   Text20(1) = ""
99 Command1(Index).Enabled = True

End Sub
Sub GRABA_MOVI_CAJA(Sali#, INGR#)
    
    SALIDA# = XVALO(Sali#)
    INGRESO# = XVALO(INGR#)
    If SALIDA# > 0 Then
       TiMovi$ = "SALECAJA"
       VARIMOVIX$ = "SALIDA"
       REFE$ = TRIM$(Text20(0))
       NUMEDOCU& = 1 + MAXNUMCAMPO&("CAJABANC", "NUMECOMP", "VARIMOVIM = 'SALIDA DE CAJA' AND TIPOMOVIM = 'SALIDA DE CAJA'")
       CueConta% = CUECOINT(Me.txtgastos)
       ImpoDeb# = Sali#
    ElseIf INGRESO# > 0 Then
       TiMovi$ = "INGRCAJA"
       VARIMOVIX$ = "INGRESO"
       REFE$ = TRIM$(Text20(1))
       NUMEDOCU& = 1 + MAXNUMCAMPO&("CAJABANC", "NUMECOMP", "VARIMOVIM = 'INGRESO MANUAL' AND TIPOMOVIM = 'INGRESO MANUAL'")
       CueConta% = CUECOINT(Me.txtingreso)
       ImpoDeb# = -1 * INGR#
    End If
    
   '
   US1% = XVALO(TXTNUMUSUARIO)
   If Abs(ImpoDeb#) < 0.005 Then
     Call MENSERR(24, "Falta Importe")
     GoTo 99
   End If
   '
   If US1% < 1 Then
     Call MENSERR(24, "Falta Usuario ")
     GoTo 99
   End If
   
   If TRIM$(REFE$) = "" Then
     Call MENSERR(24, "Falta Referencia")
     GoTo 99
   End If
   '
   FEDEP% = HOYSERVER%
   Screen.MousePointer = 11
   FETX$ = FECHATEX$(FEX)
   NRECIB$ = "TF-" + Mid$(FETX$, 7, 2) + Mid$(FETX$, 4, 2) + Mid$(FETX$, 1, 2) + "-"
   NRECIBX$ = NRECIB$ + "__"
   ORMA% = (-1)
   Call APERBASDAT("CABAN")
   Screen.MousePointer = 11
   SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK) "
   SQLX$ = SQLX$ + " WHERE CodiCom_Cor LIKE '" & NRECIBX$ & "'"
   Dim CAJABTEMP As ADODB.Recordset
   Set CAJABTEMP = READSET(SQLX$)

   On Error Resume Next
   With CAJABTEMP
30   'If .NoMatch = False Then
     .MoveFirst
     If Err < 1 Then
       On Error GoTo 0
       Do While Not .EOF
          If Val(Mid$(!CodiCom_Cor, 11)) > ORMA% Then
            ORMA% = Val(Mid$(!CodiCom_Cor, 11))
          End If
          .MoveNext
       Loop
     End If
     On Error GoTo 0
   End With
   Set CAJABTEMP = Nothing
   ORMX$ = TRIM$(Str$(ORMA% + 1))
   While Len(ORMX$) < 6: ORMX$ = "0" + ORMX$: Wend
   NRECIB$ = Left$(NRECIB$, 3) + ORMX$
   '
   CoComCor$ = NRECIB$
   CoComLar$ = NRECIB$
   NumeClie = 0
   NumeProv% = 0
   FechEmi = CVDAT(FEDEP%)
   '
   Referen$ = REFE$
   FeContab = CVDAT(FEDEP%)
   FeAcredi = CVDAT(FEDEP%)
   NumeTal = 1
   NumeCom = Val(Mid$(NRECIB$, 4))
   Observa$ = "Transferencia de Fondos"
   '
   CUDEB1% = CUECOCLI%(1) 'ver si esto esta bien
   CuentaCo% = CUDEB1%
   TiMovi = TiMovi$          ' TRANSFERENCIA DE FONDOS
   OpMovi = ""                  '
   VaMovi = VARIMOVIX$                  '
   Descrip = "TRANSFERENCIA MANUAL REALIZADA POR USUARIO: " & US1%
   'ImpoDeb# = ITRAN#
   StatInit = 0

   Call ABRECABAN
25 On Error Resume Next
   Dim CAJABANC As ADODB.Recordset
   Set CAJABANC = New ADODB.Recordset
   strsql = "SELECT * FROM CAJABANC where numecomp < 1 "
   CAJABANC.Open strsql, FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      MsgBox Err.Description
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
   With CAJABANC
     .AddNew   ' YA CONVERTIDO
     On Error Resume Next
     !Codiempr = CodiEmp%
     On Error GoTo 0
      !CodiCom_Cor = Left$(CoComCor$, 12)
      !CODICOM_LAR = Left$(CoComLar$, 24)
      !REFECOM = Left$(Referen$, 48)
      !FECHCONTA = SafeDate(FeContab)
      !fechacred = SafeDate(FeAcredi)
      !CUECONTAI = CuentaCo%
      !TipoMovim = Left$(TiMovi$, 8)
      !DescriMov = Left$(Descrip$, 128)
      !FECHEMISI = FechEmi
      !Codiempr = CodiEmp%
      !CodiSucu = CodiSuc%
      !NumeTalo = NumeTal
      !NUMECOMP = NumeCom
      !Status = StatInit
        On Error Resume Next
      !Mone_Emis = MONEMI%
      !ValAbsComp = Abs(ImpoDeb#)
      If ImpoDeb# > 0 Then
          !IDEBECOMP = ImpoDeb#
          !IHABECOMP = 0
        Else
          !IDEBECOMP = 0
          !IHABECOMP = Abs(ImpoDeb#)
      End If
      !NROCAI = TRIM$(Left$(NumCai$, 24))
      !OBSERGEN = Left$(Observa$, 64)
      !FeHoReal = Now()
      !NuAsiDiar = ""
      '
      !CUECONTAI = CuentaCo%
      USNU% = US1%
      !USERNUM = USNU%
      '
     .Update
   End With
   CAJABANC.Close
   Set CAJABANC = Nothing
   '
   CuentaCo% = CUDEB1%
   Descrip = "TRANSFERENCIA MANUAL REALIZADA POR USUARIO: " & US1%
   ImpoDeb# = (-1) * ImpoDeb#
   If YAVOLVIO% < 1 Then
     YAVOLVIO% = 1
     TiMovi = ""          ' TRANSFERENCIA DE FONDOS
     VaMovi = ""                  '
     GoTo 25
   End If
   Call COMUNI(TiMovi$ & " Completa y Verificada")
   '
99 Screen.MousePointer = 1
   Exit Sub
   
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   Call CARGA_GRILLA_CAJA
   
   Command2.Enabled = True
End Sub

Private Sub Command25_Click()
   Command25.Enabled = False
   'valido
   If MSF.Rows <= 1 Then
      MsgBox "GRILLA DE IMPRESIÓN VACIA"
      GoTo 99
   End If
   '
   FILA& = MSF.Rows
   'LE AGREGO LOS TOTALES A LA GRILLA PARA QUE SE VEA EN LA IMPRESION
    MSF.Rows = FILA& + 6
    MSF.TextMatrix(FILA& + 2, 5) = "Ingreso:": MSF.TextMatrix(FILA& + 2, 6) = Text15(0)
    MSF.TextMatrix(FILA& + 3, 5) = "Retiro:":  MSF.TextMatrix(FILA& + 3, 7) = Text15(1)
    MSF.TextMatrix(FILA& + 4, 5) = "Resultado:":  MSF.TextMatrix(FILA& + 4, 7) = Text15(2)

   Set GRILLA_IMPRESION = MSF
   '
   PREIMPRE.txtEncabezado2.Text = ""
   '
   ENCABEZADO1$ = "Reporte de Movimientos de Caja"
   
   PREIMPRE.txtEncabezado1.Text = ENCABEZADO1$
   PREIMPRE.txtPiePagina = "Fin de Reporte"
   PREIMPRE.Combo1.Text = "Vertical"
   PREIMPRE.Show 1
   '
   'ESTO ES PARA VOLVER A CORTAR LA GRILLA DONDE ESTABA.
   MSF.Rows = FILA&
99 Command25.Enabled = True

End Sub



Private Sub Command4_Click()
    Command4.Enabled = False
    If DERACCE%("CAMBUSUAR", "VISUALIZAR MOVIMIENTOS DE CAJA DE OTROS USUARIOS") < 2 Then GoTo 99
    Screen.MousePointer = 11
    Call BLOQARCH("FLEXUSR")
    FUSER$ = LOADFILE$("USER.DAT")
    JJI& = 0: KKI% = 0
    For i% = 1 To Len(FUSER$) Step 64
       F1$ = Mid$(FUSER$, i%, 64)
       If Asc(F1$) = 0 Then GoTo 9
       NOM$ = Left$(F1$, 30)
       If Left$(NOM$, 1) = "@" Then NOM$ = DESENCRI$(Mid$(F1$, 2, 29))
       '
       KKI% = KKI% + 1
       X$ = REGBLAN$("FLEXUSR")
       Call REPLA$(X$, MKI$(KKI%), 1, 2)
       Call REPLA$(X$, NOM$, 3, 30)
       JJI& = JJI& + 1
       Call GRAREG("FLEXUSR", X$, JJI&)
    Next i%
    '
9   Call SETULTREG("FLEXUSR", JJI&)
10  Call CIERRARCH("FLEXUSR")
    '
    Call SELECHO("FLEXUSR")
    Screen.MousePointer = 1
    If VALACT1$("FLEXUSR") = "" Then GoTo 99
    '
    USEL% = XVALO(TRIM$(VALACT1$("FLEXUSR")))
    DEUSER$ = VALACT2$("FLEXUSR")
    Me.TXTNUMUSUARIO = USEL%
    Me.TXTDESCRIUSUARIO = DEUSER$
99  Command4.Enabled = True
End Sub

Private Sub Form_Load()
    Call APLICAR_SKIN(Me, App.Path & "\SKINS\SteelBlue.skn")
    Call ENCABEZA_GRILLA
    Desde.Value = Date
    Hasta.Value = Date
    'Call CARGA_GRILLA_CAJA
    USNX% = USNBR% Mod 100
    Me.TXTNUMUSUARIO = USNX%
    Me.txtgastos = TRIM$(CONTROL$("RECISTA", "CUECOREC"))
    Me.txtingreso = TRIM$(CONTROL$("", "FORAJUSA"))
'    Call CARGAVECTOR_MES
    '
End Sub


Sub CARGA_GRILLA_CAJA()
   '
   MSF.Rows = 1
' Call LIMPIASUBTOTALES
   '
   Screen.MousePointer = 11
   DES% = CVINT(Desde.Value)
   HAS% = CVINT(Hasta.Value)
   DESDEX = FETEXCOR1$(DES%)
   HASTAX = FETEXCOR1$(HAS%)
   FILA& = 0
   'CARGA EL ENCABEZADO DEL TRANSPORTE
   SQLX$ = "select * from cajabanc where "
   CONDI$ = " (tipomovim = 'rcob' OR TipoMovim = 'INGRCAJA' OR TipoMovim = 'SALECAJA' OR TIPOMOVIM = 'TRAFON')"
   CONDI$ = CONDI$ + "  AND FECHEMISI < '" & DESDEX & "'"
   If XVALO(TXTNUMUSUARIO) > 0 Then CONDI$ = CONDI$ + "  AND USERNUM = " & XVALO(Me.TXTNUMUSUARIO)
   If CantRegistros&("CAJABANC", CONDI$, "NOMESS") < 1 Then GoTo 10
   
   SQLX$ = "SELECT SUM(idebecomp) AS DEBEX, SUM(ihabecomp)AS HABEX"
   SQLX$ = SQLX$ + " FROM CAJABANC WITH(NOLOCK) "
   SQLX$ = SQLX$ + " WHERE " + CONDI$
'   SQLX$ = SQLX$ + " group by CUENTA_CONT"

   Set RECITMP0 = FLEXCONN.Execute(SQLX$)
   With RECITMP0
      Do While Not .EOF
           RETIRO# = XVALO(!DEBEX)
           SUPAGO# = XVALO(!HABEX)
           SUINGREACUM# = SUINGREACUM# + SUPAGO#
           RETIROACUM# = RETIROACUM# + RETIRO#
           '
           FILA& = FILA& + 1
           TXTCANTIFILAS = FILA&: TXTCANTIFILAS.Refresh
           MSF.Rows = FILA& + 1
           MSF.TextMatrix(FILA&, 1) = "-------------------"
           MSF.TextMatrix(FILA&, 2) = "-------------------"
           MSF.TextMatrix(FILA&, 3) = "-------------------"
           MSF.TextMatrix(FILA&, 4) = "-------------------"
           MSF.TextMatrix(FILA&, 5) = "Transporte................"
           'MSF.TextMatrix(FILA&, 6) = "-------------------"
           MSF.TextMatrix(FILA&, 6) = TRIM$(FORMATNUM$(SUPAGO#, "F12.2"))
           MSF.TextMatrix(FILA&, 7) = TRIM$(FORMATNUM$(RETIRO#, "F12.2"))
           
           .MoveNext
      Loop
    End With
    RECITMP0.Close
    Set RECITMP0 = Nothing
    FILA& = FILA& + 1
    '
    'BLANQEO POR QUE PIDIERON QUE SOLO ACUMULE EL TOTAL DEL RANGO SELECCIONADO.
'    SUINGREACUM# = 0
'    RETIROACUM# = 0

10 CONDI$ = " (tipomovim = 'rcob' OR TipoMovim = 'INGRCAJA' OR TipoMovim = 'SALECAJA' OR TIPOMOVIM = 'TRAFON')"
   CONDI$ = CONDI$ + " AND FECHEMISI >= '" & DESDEX & "' AND FECHEMISI <= '" & HASTAX & "'"
   If XVALO(TXTNUMUSUARIO) > 0 Then CONDI$ = CONDI$ + "  AND USERNUM = " & XVALO(Me.TXTNUMUSUARIO)
   '
   If CantRegistros&("CAJABANC", CONDI$, "NOMESS") < 1 Then GoTo 90
   '
   'MSF.Redraw = False
   '
   Dim CLS As New clientes
   CLS.cargarvector
   SQLX$ = "SELECT IDEBECOMP,TipoMovim, IHABECOMP,FECHEMISI,DescriMov,NUCLIEN,CODICOM_LAR,REFECOM  FROM CAJABANC WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE " + CONDI$
   SQLX$ = SQLX$ + " ORDER BY FeHoReal"
   Set RECITMP = FLEXCONN.Execute(SQLX$)
   With RECITMP
      Do While Not .EOF
           RETIRO# = XVALO(!IDEBECOMP)
           SUINGRE# = XVALO(!IHABECOMP)
           FEMI% = CVINT(!FECHEMISI)
           REFE$ = SAFETEXT$(!DescriMov)
           If SAFETEXT$(!TipoMovim) = "INGRCAJA" Or SAFETEXT$(!TipoMovim) = "SALECAJA" Then REFE$ = SAFETEXT$(!REFECOM)
'           CUO_ORIGINALX# = XVALO(!CUOTA_ORIGINAL)
           LGJO& = XVALO(!NUCLIEN)
'           Debug.Print LGJO&
'           Debug.Print CLS.NOMBRE(LGJO&)
           DOCUM$ = SAFETEXT$(!CODICOM_LAR)
           SUINGREACUM# = SUINGREACUM# + SUINGRE#
           RETIROACUM# = RETIROACUM# + RETIRO#
           '
           FILA& = FILA& + 1
           TXTCANTIFILAS = FILA&: TXTCANTIFILAS.Refresh
           MSF.Rows = FILA& + 1
           MSF.TextMatrix(FILA&, 1) = FECHATEX$(FEMI%)
           MSF.TextMatrix(FILA&, 2) = DOCUM$
           MSF.TextMatrix(FILA&, 3) = REFE$
           MSF.TextMatrix(FILA&, 4) = TRIM$(Str$(LGJO&))
           MSF.TextMatrix(FILA&, 5) = CLS.nombre(LGJO&)
'           MSF.TextMatrix(FILA&, 6) = TRIM$(FORMATNUM$(CUO_ORIGINALX#, "F12.2"))
           MSF.TextMatrix(FILA&, 6) = TRIM$(FORMATNUM$(SUINGRE#, "F12.2"))
           MSF.TextMatrix(FILA&, 7) = TRIM$(FORMATNUM$(RETIRO#, "F12.2"))
'           Call ResaltarFila(MSF, FILA&)
           .MoveNext
      Loop
    End With
    RECITMP.Close
    Set RECITMP = Nothing
    '
90  'Call ENCABEZA_GRILLA

    Call SUMATOTAL(SUINGREACUM#, RETIROACUM#)
'    MSF.Redraw = True

    'CARGO A PIE DE LA GRILLA EL TOTOL
'    MSF.Rows = FILA& + 6
'    MSF.TextMatrix(FILA& + 2, 5) = "Ingreso:": Call COLOREAR_GRILLA_SOLO_UNA_AFILA(CAJAINS.MSF, &HFCE4A7, FILA& + 2)
'    MSF.TextMatrix(FILA& + 3, 5) = "Retiro:": Call COLOREAR_GRILLA_SOLO_UNA_AFILA(CAJAINS.MSF, &HFCE4A7, FILA& + 3)
    
'    MSF.TextMatrix(FILA& + 4, 5) = "Resultado:": Call COLOREAR_GRILLA_SOLO_UNA_AFILA(CAJAINS.MSF, &HFCE4A7, FILA& + 4)
    '
   'MSF.TextMatrix(FILA& + 2, 6) = Text15(0)
   'MSF.TextMatrix(FILA& + 3, 7) = Text15(1)
    'SI ES POSITIVO VA DEL LADO DE INGRESO SI ES NEGATIVO VA DEL LADO RETIRO
   ' POS% = 7
    'If XVALO(Text15(2)) > 0 Then POS% = 6
    'MSF.TextMatrix(FILA& + 4, 6) = Text15(1)
'    MSF.TextMatrix(FILA& + 4, 7) = Text15(2)
    '
    Screen.MousePointer = 1

End Sub
Sub SUMATOTAL(SUPAGOACUM#, RETIROACUM#)
    Text15(0) = FORMATNUM$(SUPAGOACUM#, "F14.2")
    Text15(1) = FORMATNUM$(RETIROACUM#, "F14.2")
    Text15(2) = FORMATNUM$(SUPAGOACUM# - RETIROACUM#, "F14.2")
    
End Sub
Sub ENCABEZA_GRILLA()
    Call CARGA_ENCABEZADO(MSF, "Fecha Em./D8;Documento/A18;Referencia/A32;Lgjo/f7;Nombre y Apellido/A24;Ingreso/F12.2;Retiro/F12.2", Me)
    '
'    MSF.MergeCells = flexMergeRestrictColumns
'    MSF.MergeCol(5) = True

End Sub

Sub LIMPIASUBTOTALES()

   For i% = 0 To 2
      Text15(i%) = ""
   Next i%
End Sub


Private Sub mnucajadetalle_Click()
  CAJADET.Show 1
End Sub

Private Sub mnuTRansfxUsuario_Click()
     
   FEX = HOY(HOS$): Call REPLA(VDEF$, MKI$(FEX), 1, 2)
   NUMUS% = USNBR% Mod 100
   Call REPLA(VDEF$, SAFETEXT$(NUMUS%), 11, 4)
   '
   OBJETOX$ = "TRANSFE-USER"
   If NUMUS% = 1 Then OBJETOX$ = "Transferencias D"
10 DTP$ = OBJPLA$(OBJETOX$, VDEF$)
   If TRIM$(DTP$) = "" Then Exit Sub
   VDEF$ = DTP$
   
   FEX = CVI(Left(VDEF$, 2))
   ImpoDeb# = CVD(Mid$(VDEF$, 3, 8))
   US1% = XVALO(Mid$(VDEF$, 11, 4))
   US2% = XVALO(Mid$(VDEF$, 15, 4))
   referencia$ = Mid$(VDEF$, 19, 40)
   '
   
   If FEX < 1 Then
     Call MENSERR(24, "Fecha no Válida")
     GoTo 10
   End If
   '
   If ImpoDeb# < 0.005 Then
     Call MENSERR(24, "Falta Importe")
     GoTo 10
   End If
   '
   If US1% < 1 Then
     Call MENSERR(24, "Falta Usuario Desde")
     GoTo 10
   End If
   If US2% < 1 Then
     Call MENSERR(24, "Falta Usuario Hasta")
     GoTo 10
   End If
   
   If TRIM$(referencia) = "" Then
     Call MENSERR(24, "Falta Referencia")
     GoTo 10
   End If
   
   REFE$ = TRIM$(Mid$(OBX$, 35, 40)) + " " + FECHATEX$(FEX)
   '
   FEDEP% = FEX
   Screen.MousePointer = 11
   FETX$ = FECHATEX$(FEX)
   NRECIB$ = "TF-" + Mid$(FETX$, 7, 2) + Mid$(FETX$, 4, 2) + Mid$(FETX$, 1, 2) + "-"
   NRECIBX$ = NRECIB$ + "__"
   ORMA% = (-1)
   Call APERBASDAT("CABAN")
   Screen.MousePointer = 11
   SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK) "
   SQLX$ = SQLX$ + " WHERE CodiCom_Cor LIKE '" & NRECIBX$ & "'"
   Dim CAJABTEMP As ADODB.Recordset
   Set CAJABTEMP = READSET(SQLX$)

   On Error Resume Next
   With CAJABTEMP
30   'If .NoMatch = False Then
     .MoveFirst
     If Err < 1 Then
       On Error GoTo 0
       Do While Not .EOF
          If Val(Mid$(!CodiCom_Cor, 11)) > ORMA% Then
            ORMA% = Val(Mid$(!CodiCom_Cor, 11))
          End If
          .MoveNext
       Loop
     End If
     On Error GoTo 0
   End With
   Set CAJABTEMP = Nothing
   ORMX$ = TRIM$(Str$(ORMA% + 1))
   While Len(ORMX$) < 2: ORMX$ = "0" + ORMX$: Wend
   NRECIB$ = NRECIB$ + Left$(ORMX$, 2)
   '
   CoComCor$ = NRECIB$
   CoComLar$ = NRECIB$
   NumeClie = 0
   NumeProv% = 0
   FechEmi = CVDAT(FEDEP%)
   '
   Referen$ = REFE$
   FeContab = CVDAT(FEDEP%)
   FeAcredi = CVDAT(FEDEP%)
   NumeTal = Val(Mid$(NRECIB$, 11, 2))
   NumeCom = Val(Mid$(NRECIB$, 4, 6))
   Observa$ = "Transferencia de Fondos"
   '
   CUDEB1% = CUECOCLI%(1) 'ver si esto esta bien
   CuentaCo% = CUDEB1%
   TiMovi = "TRAFON"            ' TRANSFERENCIA DE FONDOS
   OpMovi = ""                  '
   VaMovi = ""                  '
   Descrip = "DE USUARIO NRO.: " & US1%
   'ImpoDeb# = ITRAN#
   StatInit = 0

   Call ABRECABAN
25 On Error Resume Next
   Dim CAJABANC As ADODB.Recordset
   Set CAJABANC = New ADODB.Recordset

   
   strsql = "SELECT * FROM CAJABANC "
   CAJABANC.Open strsql, FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      MsgBox Err.Description
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
   With CAJABANC
     .AddNew   ' YA CONVERTIDO
     On Error Resume Next
     !Codiempr = CodiEmp%
     On Error GoTo 0
      !CodiCom_Cor = Left$(CoComCor$, 12)
      !CODICOM_LAR = Left$(CoComLar$, 24)
      !REFECOM = Left$(Referen$, 48)
      !FECHCONTA = SafeDate(FeContab)
      !fechacred = SafeDate(FeAcredi)
      !CUECONTAI = CuentaCo%
      !TipoMovim = Left$(TiMovi$, 8)
      !DescriMov = Left$(Descrip$, 128)
      !FECHEMISI = FechEmi
      !Codiempr = CodiEmp%
      !CodiSucu = CodiSuc%
      !NumeTalo = NumeTal
      !NUMECOMP = NumeCom
      !Status = StatInit
        On Error Resume Next
      !Mone_Emis = MONEMI%
      !ValAbsComp = Abs(ImpoDeb#)
      If ImpoDeb# > 0 Then
          !IDEBECOMP = ImpoDeb#
          !IHABECOMP = 0
        Else
          !IDEBECOMP = 0
          !IHABECOMP = Abs(ImpoDeb#)
      End If
      !NROCAI = TRIM$(Left$(NumCai$, 24))
      !OBSERGEN = Left$(Observa$, 64)
      !FeHoReal = Now()
      !NuAsiDiar = ""
      '
      USNU% = US1%
      !USERNUM = USNU%
      '
     .Update
   End With
   CAJABANC.Close
   Set CAJABANC = Nothing
   '
   CuentaCo% = CUDEB1%
   Descrip = "A USUARIO NRO.: " & US2%
   ImpoDeb# = (-1) * ImpoDeb#
   US1% = US2%
   If YAVOLVIO% < 1 Then
     YAVOLVIO% = 1
     GoTo 25
   End If
   Call COMUNI("Transferencia Completa y Verificada")
   '
   Screen.MousePointer = 1
End Sub





Private Sub MSF_DblClick()
        REG& = MSF.MouseRow
        If REG& < 1 Then Exit Sub
        
        TIDOCUM$ = MSF.TextMatrix(REG&, 2)
        If InStr(TIDOCUM$, "RB-") Then TIDOCUM$ = "Recibo"
        FE% = FECHANUM(MSF.TextMatrix(REG&, 1))
        NUDOCU$ = " " + CStr(XVALO(Mid$(MSF.TextMatrix(REG&, 2), 11, 8))) + " "
        NUDOCU1$ = Str$(XVALO(NUDOCU$))
        NUDOCU2$ = NUDOCU$
        
        TIPOCOM% = 3
        CONDI$ = "NUMECOMP= " & XVALO(NUDOCU$) & " AND fechemisi = '" & FETEXCOR1$(FE%) & "' and TIPOMOVIM = 'RCOB'"
        
        FE% = CVINT(VALOBADA("CAJABANC", "FeHoReal", CONDI$, "NOMESS"))
        NC = XVALO(VALOBADA("CAJABANC", "NUCLIEN", CONDI$, "NOMESS"))
        NOMB$ = rasocli$(NC)
        
        
        If TIDOCUM$ = "Recibo" Then VARCONTROL$ = NOMB$
        DOCUNU& = DOCID&(TIDOCUM$, NUDOCU1$, NUDOCU2$, VARCONTROL$)
        If DOCUNU& < 1 Then
         Call COMUNI("Documento No Encontrado")
        Else
          OPPXX$ = ""
        If NUDOCU$ = "00000032" Then OPPXX$ = "RECUP"
          Call MUESTRADOC(DOCUNU&, OPPXX$, TIDOCUM$, NUDOCU1$, NUDOCU2$)
        End If

End Sub


Private Sub Text2_LOSTFOCUS(Index As Integer)
    Text2(Index) = FORMATNUM$(XVALO(Text2(Index)), "F14.2")
End Sub

Private Sub txtgastos_Change()
 Me.Text16(4) = DENOCUE$(CUECOINT(Me.txtgastos))
End Sub

Private Sub txtgastos_DblClick()
   
   
           USNX% = USNBR% Mod 100
           If USNX% <> 1 Then
              Call COMUNI("Solo el Usuario Supervisor Puede Realizar Cambios en las Cuentas Contables")
              Exit Sub
           End If
        
           CuentaContable$ = TRIM$(CONTROL$("RECISTA", "CUECOREC"))
    
            Call SELECHO("ECUECON")
            
            CueConta$ = TRIM$(VALACT1$("ECUECON"))
            If CueConta$ <> "" Then
                Me.txtgastos = CueConta$
                Call GRACONTROL("RECISTA", "CUECOREC", CueConta$)
                Call COMUNI("Se Ha Actualizdo la Cuenta Contable Para Gastos Varios")
            End If
End Sub


Private Sub txtingreso_Change()
    Me.Text16(5) = DENOCUE$(CUECOINT(Me.txtingreso))
End Sub


Private Sub txtingreso_DblClick()
            CuentaContable$ = TRIM$(CONTROL$("", "FORAJUSA"))
            Call SELECHO("ECUECON")
            CueConta$ = TRIM$(VALACT1$("ECUECON"))
            If CueConta$ <> "" Then
                Call GRACONTROL("", "FORAJUSA", CueConta$)
                Me.txtingreso = CueConta$
                Call COMUNI("Se Ha Actualizdo la Cuenta Contable Para Gastos Varios")
            End If

End Sub


