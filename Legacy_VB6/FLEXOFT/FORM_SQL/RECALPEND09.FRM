VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form RECALPEND09 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Detalle de Composición de Deuda"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11655
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   11655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00D8BD05&
      Height          =   540
      Left            =   120
      TabIndex        =   18
      Top             =   7680
      Width           =   10515
      Begin VB.Label Label19 
         BackColor       =   &H00D8BD05&
         BackStyle       =   0  'Transparent
         Caption         =   "Total $ + U$S = "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3720
         TabIndex        =   34
         Top             =   240
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D8BD05&
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4800
         TabIndex        =   33
         Top             =   240
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5115
         TabIndex        =   32
         Top             =   165
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   7080
         TabIndex        =   28
         Top             =   165
         Visible         =   0   'False
         Width           =   1530
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D8BD05&
         BackStyle       =   0  'Transparent
         Caption         =   "u$s"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6720
         TabIndex        =   27
         Top             =   220
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D8BD05&
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   8640
         TabIndex        =   26
         Top             =   225
         Width           =   180
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   8880
         TabIndex        =   20
         Top             =   165
         Width           =   1530
      End
      Begin VB.Label Label6 
         BackColor       =   &H00D8BD05&
         BackStyle       =   0  'Transparent
         Caption         =   "Tot. Selección:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7500
         TabIndex        =   19
         Top             =   220
         Width           =   1320
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C000&
      Height          =   540
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   10515
      Begin VB.Label Label2 
         BackColor       =   &H00D8BD05&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2295
         TabIndex        =   11
         Top             =   240
         Width           =   3540
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D8BD05&
         BackStyle       =   0  'Transparent
         Caption         =   "Rango Seleccionado:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   1980
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      Height          =   820
      Left            =   120
      TabIndex        =   7
      Top             =   3720
      Width           =   10515
      Begin VB.ComboBox CMBFILTRO 
         Height          =   315
         Left            =   3000
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   400
         Width           =   2055
      End
      Begin VB.TextBox TXTBUSCAR 
         Height          =   315
         Left            =   240
         TabIndex        =   12
         Top             =   400
         Width           =   2535
      End
      Begin VB.Label Label16 
         BackColor       =   &H00D8BD05&
         BackStyle       =   0  'Transparent
         Caption         =   "Total $ + U$S"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5040
         TabIndex        =   31
         Top             =   210
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D8BD05&
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5020
         TabIndex        =   30
         Top             =   450
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
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
         Left            =   5340
         TabIndex        =   29
         Top             =   405
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D8BD05&
         BackStyle       =   0  'Transparent
         Caption         =   "u$s"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6960
         TabIndex        =   25
         Top             =   480
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
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
         Left            =   7320
         TabIndex        =   24
         Top             =   405
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
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
         Left            =   9000
         TabIndex        =   22
         Top             =   405
         Width           =   1455
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D8BD05&
         BackStyle       =   0  'Transparent
         Caption         =   "Campo:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2760
         TabIndex        =   15
         Top             =   210
         Width           =   705
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D8BD05&
         BackStyle       =   0  'Transparent
         Caption         =   "Filtrar:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   0
         TabIndex        =   14
         Top             =   210
         Width           =   700
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D8BD05&
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   8680
         TabIndex        =   8
         Top             =   450
         Width           =   300
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   225
      Index           =   0
      Left            =   1470
      ScaleHeight     =   165
      ScaleWidth      =   1110
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00800000&
      Height          =   8220
      Left            =   10680
      ScaleHeight     =   8160
      ScaleWidth      =   1845
      TabIndex        =   0
      Top             =   0
      Width           =   1905
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   120
         Picture         =   "RECALPEND09.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Imprime Información de Grilla Superior"
         Top             =   2880
         Width           =   680
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
         Height          =   650
         Left            =   120
         Picture         =   "RECALPEND09.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Imprime Información de Grilla Inferior"
         Top             =   6720
         Width           =   680
      End
      Begin VB.CommandButton BOTEXIT 
         Caption         =   "Cancel"
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
         Left            =   100
         Picture         =   "RECALPEND09.frx":0CD4
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Cancelar la Aplicación"
         Top             =   105
         Width           =   680
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   130
         ScaleHeight     =   75
         ScaleWidth      =   540
         TabIndex        =   1
         Top             =   7440
         Width           =   600
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   2
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -240
         Picture         =   "RECALPEND09.frx":0FDE
         Top             =   7680
         Width           =   1515
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "RECALPEND09.frx":2F00
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "RECALPEND09.frx":3134
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   3045
      Left            =   120
      TabIndex        =   16
      ToolTipText     =   "Un Click Sobre Encabezado Para Ordenar/Doble Click Para Pasar a Grilla Inferior/Click Boton Derecho Imagen Digital"
      Top             =   600
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   5371
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin MSFlexGridLib.MSFlexGrid MSF2 
      Height          =   2925
      Left            =   120
      TabIndex        =   17
      ToolTipText     =   "Un Click Sobre Encabezado Para Ordenar/Doble Click Para Pasar a Grilla Superior/Click Boton Derecho Imagen Digital"
      Top             =   4680
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   5159
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   8438015
      BackColorSel    =   33023
      BackColorBkg    =   16777215
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin MSFlexGridLib.MSFlexGrid MSF_2 
      Height          =   3045
      Left            =   120
      TabIndex        =   23
      ToolTipText     =   "Un Click Sobre Encabezado Para Ordenar/Doble Click Para Pasar a Grilla Inferior/Click Boton Derecho Imagen Digital"
      Top             =   600
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   5371
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   8421631
      BackColorBkg    =   12632256
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "RECALPEND09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub BOTEXIT_Click()
  Unload Me
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command1_Click()
   Command1.Enabled = False
   If MSF1.Rows <= 1 Then
     Call COMUNI("Sin Información en Grilla Inferior")
     GoTo 99
   End If
   '
   Call BORRAREGISTROS("COMPPEDI", "NroDocumento >=0 ", REGAF&, "")
   '
   SQLX$ = "SELECT * FROM COMPPEDI ORDER BY NroCuenta"
   Dim RS As ADODB.Recordset
   Set RS = New ADODB.Recordset
25 On Error Resume Next
   RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
   '
   With RS
    For i& = 1 To MSF1.Rows - 1
      .AddNew
      !NroDocumento = XVALO(MSF1.TextMatrix(i&, 1))
      !NroCuenta = XVALO(MSF1.TextMatrix(i&, 2))
      !razonSocial = MSF1.TextMatrix(i&, 3)
      !FECHENT = CVDAT(CVINT(MSF1.TextMatrix(i&, 4)))
      !IMPORTE = XVALO(MSF1.TextMatrix(i&, 5))
      !Moneda = MSF1.TextMatrix(i&, 6)
      .Update
    Next i&
   End With
   
   FILTX$ = UCase$(RECALPEND09.Caption) & ";" & TRIM$(Label2) & ";" & FECHATEX$(HOY(HO$))
   Call STDFORM("REC-JYM", FILTX$)
   '
99 Command1.Enabled = True

End Sub

Private Sub Command3_Click()
   '
   Command3.Enabled = False
   If msf2.Rows <= 1 Then
     Call COMUNI("Sin Información en Grilla Inferior")
     GoTo 99
   End If
   '
   Call BORRAREGISTROS("COMPPEDI", "NroDocumento >=0 ", REGAF&, "")
   '
   SQLX$ = "SELECT * FROM COMPPEDI order by NroCuenta"
   Dim RS As ADODB.Recordset
   Set RS = New ADODB.Recordset
25 On Error Resume Next
   RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
   '
   With RS
    For i& = 1 To msf2.Rows - 1
      .AddNew
      !NroDocumento = XVALO(msf2.TextMatrix(i&, 1))
      !NroCuenta = XVALO(msf2.TextMatrix(i&, 2))
      !razonSocial = msf2.TextMatrix(i&, 3)
      !FECHENT = CVDAT(CVINT(msf2.TextMatrix(i&, 4)))
      !IMPORTE = XVALO(msf2.TextMatrix(i&, 5))
      !Moneda = msf2.TextMatrix(i&, 6)
      .Update
    Next i&
   End With
   
   FILTX$ = UCase$(RECALPEND09.Caption) & ";" & TRIM$(Label2) & ";" & FECHATEX$(HOY(HO$))
   Call STDFORM("REC-JYM", FILTX$)
'    If STDSCREEN.APRO <> "OK" Then
'       If OPX% = 2 Then GoTo 15
'       Exit Sub
'    End If
99 Command3.Enabled = True

End Sub

Private Sub Form_Load()
   Call ABRECONEXION
   Call CREACAMPO("", "COMPPEDI", "NroDocumento", 4, 0, 1)
   Call CREACAMPO("", "COMPPEDI", "NroCuenta", 4, 0, 1)
   Call CREACAMPO("", "COMPPEDI", "RazonSocial", 10, 64)
   Call CREACAMPO("", "COMPPEDI", "FechEnt", 8, 0)
   Call CREACAMPO("", "COMPPEDI", "Importe", 7, 0)
   Call CREACAMPO("", "COMPPEDI", "MONEDA", 10, 8)

End Sub

Private Sub Label13_Change()
   Call ACTUTOTA2
End Sub

Private Sub Label5_Click()
   Call ACTUTOTA2
End Sub

Private Sub Label8_Change()
    Call ACTUTOTA1
End Sub

Private Sub Label9_Change()
    Call ACTUTOTA1
End Sub
Sub ACTUTOTA1()
   Static COTIHOY#
   If COTIHOY# = 0 Then
      COTIHOY# = TICAMONE(2)
   End If
   If COTIHOY# = 0 Then COTIHOY# = 1
   Label14 = FORMATNUM$((XVALO(Label9) * COTIHOY#) + XVALO(Label8), "F12.2")
   Label17 = FORMATNUM$((XVALO(Label13) * COTIHOY#) + XVALO(Label5), "F12.2")
End Sub
Sub ACTUTOTA2()
   Static COTIHOY#
   If COTIHOY# = 0 Then
      COTIHOY# = TICAMONE(2)
   End If
   If COTIHOY# = 0 Then COTIHOY# = 1
   Label17 = FORMATNUM$((XVALO(Label13) * COTIHOY#) - XVALO(Label5), "F12.2")
   Label14 = FORMATNUM$((XVALO(Label9) * COTIHOY#) + XVALO(Label8), "F12.2")
End Sub

Private Sub MSF1_Click()
    Static MODO  As Boolean
    If (MSF1.MouseRow = 0) Then
          
        ' Ordena en forma ascendente
        If MODO Then
            MSF1.COL = MSF1.MouseCol
            MSF1.Sort = 2
            MODO = False
        ' Ordena en forma descendente
        Else
            MSF1.COL = MSF1.MouseCol
            MSF1.Sort = 1
            MODO = True
        End If
    End If
End Sub
Private Sub MSF_2_Click()
    Static MODO  As Boolean
    If (MSF_2.MouseRow = 0) Then
          
        ' Ordena en forma ascendente
        If MODO Then
            MSF_2.COL = MSF_2.MouseCol
            MSF_2.Sort = 2
            MODO = False
        ' Ordena en forma descendente
        Else
            MSF_2.COL = MSF_2.MouseCol
            MSF_2.Sort = 1
            MODO = True
        End If
    End If
End Sub
Private Sub MSF1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
    '
    i& = MSF1.MouseRow
    J& = MSF1.MouseCol
    If i& = 0 Or i& > MSF1.Rows Then Exit Sub
    '
    On Error Resume Next
    MSF1.Row = i&
    MSF1.COL = J&
    On Error Resume Next

    NPX& = TRIM$(MSF1.TextMatrix(i&, 1))
    If NPX& < 1 Then
        Exit Sub
    End If

    If Me.Caption = "Pedidos" Then
        '
        NUDOCU$ = TRIM$(Str$(NPX&))
        TIDOCUM$ = "Pedido de Cliente"
        While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
        NUDOCU$ = NUDOCU$ + " "
        Screen.MousePointer = 11
        '
        'BUSCA EL DOCUMENTO EN LA TABLA SQL CON LA NUEVA FUNCION
        DOCUNU& = DOCID&(TIDOCUM$, Str$(NPX&), NUDOCU$)
'        If DOCUNU& < 1 Then
'             Screen.MousePointer = 1
'             Call COMUNI("No se Encontró el Documento Impreso\Correspondiente a la " & DOCUMENTO_ELEGIDO$ & " .\  \Si se Trata de un Pedido Pendiente de Entrega\Puede Refrescarlo con la Opción\'Modificación de Pedido en Curso'.")
'             On Error Resume Next
'             OPPED07.Command46.SetFocus
'             OPPED07.Show 1
'             If Err Then
'              On Error GoTo 0
'             End If
'             Exit Sub
'        End If
        Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, Str$(NPX&), NUDOCU$)
    Else
        TIDOCUM$ = "Nota de Pedido"
        NUDOCU$ = TRIM$(Str$(NPX&))
        DOCUNU& = 0
        '
        Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, NUDOCU$)
    End If
    Screen.MousePointer = 1
  End If
End Sub
Private Sub MSF_2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
    '
    i& = MSF_2.MouseRow
    J& = MSF_2.MouseCol
    If i& = 0 Or i& > MSF_2.Rows Then Exit Sub
    '
    On Error Resume Next
    MSF_2.Row = i&
    MSF_2.COL = J&
    On Error Resume Next

    NPX& = TRIM$(MSF_2.TextMatrix(i&, 1))
    If NPX& < 1 Then
        Exit Sub
    End If

    If Me.Caption = "Pedidos" Then
        '
        NUDOCU$ = TRIM$(Str$(NPX&))
        TIDOCUM$ = "Pedido de Cliente"
        While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
        NUDOCU$ = NUDOCU$ + " "
        Screen.MousePointer = 11
        '
        'BUSCA EL DOCUMENTO EN LA TABLA SQL CON LA NUEVA FUNCION
        DOCUNU& = DOCID&(TIDOCUM$, Str$(NPX&), NUDOCU$)
'        If DOCUNU& < 1 Then
'             Screen.MousePointer = 1
'             Call COMUNI("No se Encontró el Documento Impreso\Correspondiente a la " & DOCUMENTO_ELEGIDO$ & " .\  \Si se Trata de un Pedido Pendiente de Entrega\Puede Refrescarlo con la Opción\'Modificación de Pedido en Curso'.")
'             On Error Resume Next
'             OPPED07.Command46.SetFocus
'             OPPED07.Show 1
'             If Err Then
'              On Error GoTo 0
'             End If
'             Exit Sub
'        End If
        Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, Str$(NPX&), NUDOCU$)
    Else
        TIDOCUM$ = "Nota de Pedido"
        NUDOCU$ = TRIM$(Str$(NPX&))
        DOCUNU& = 0
        '
        Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, NUDOCU$)
    End If
    Screen.MousePointer = 1
  End If
End Sub

Private Sub msf2_Click()
    Static MODO  As Boolean
    If (msf2.MouseRow = 0) Then
          
        ' Ordena en forma ascendente
        If MODO Then
            msf2.COL = msf2.MouseCol
            msf2.Sort = 2
            MODO = False
        ' Ordena en forma descendente
        Else
            msf2.COL = msf2.MouseCol
            msf2.Sort = 1
            MODO = True
        End If
    End If

End Sub

Private Sub MSF2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
    '
    i& = msf2.MouseRow
    J& = msf2.MouseCol
    If i& = 0 Or i& > msf2.Rows Then Exit Sub
    '
    On Error Resume Next
    msf2.Row = i&
    msf2.COL = J&
    On Error Resume Next

    NPX& = TRIM$(msf2.TextMatrix(i&, 1))
    If NPX& < 1 Then
        Exit Sub
    End If

    If Me.Caption = "Pedidos" Then
        '
        NUDOCU$ = TRIM$(Str$(NPX&))
        TIDOCUM$ = "Pedido de Cliente"
        While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
        NUDOCU$ = NUDOCU$ + " "
        Screen.MousePointer = 11
        '
        'BUSCA EL DOCUMENTO EN LA TABLA SQL CON LA NUEVA FUNCION
        DOCUNU& = DOCID&(TIDOCUM$, Str$(NPX&), NUDOCU$)
'        If DOCUNU& < 1 Then
'             Screen.MousePointer = 1
'             Call COMUNI("No se Encontró el Documento Impreso\Correspondiente a la " & DOCUMENTO_ELEGIDO$ & " .\  \Si se Trata de un Pedido Pendiente de Entrega\Puede Refrescarlo con la Opción\'Modificación de Pedido en Curso'.")
'             On Error Resume Next
'             OPPED07.Command46.SetFocus
'             OPPED07.Show 1
'             If Err Then
'              On Error GoTo 0
'             End If
'             Exit Sub
'        End If
        Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, Str$(NPX&), NUDOCU$)
    Else
        TIDOCUM$ = "Nota de Pedido"
        NUDOCU$ = TRIM$(Str$(NPX&))
        DOCUNU& = 0
        '
        Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, NUDOCU$)
    End If
    Screen.MousePointer = 1
  End If
End Sub


Private Sub MSF1_DblClick()
   If (MSF1.MouseRow = 0) Then Exit Sub
   REG& = MSF1.Row
   
   Static COTIHOY#
   If COTIHOY# = 0 Then
      COTIHOY# = TICAMONE(2)
   End If
   If COTIHOY# = 0 Then COTIHOY# = 1
   'If REG& < 1 Then Exit Sub
   msf2.Rows = msf2.Rows + 1
   msf2.TextMatrix(msf2.Rows - 1, 1) = MSF1.TextMatrix(REG&, 1)
   msf2.TextMatrix(msf2.Rows - 1, 2) = MSF1.TextMatrix(REG&, 2)
   msf2.TextMatrix(msf2.Rows - 1, 3) = MSF1.TextMatrix(REG&, 3)
   msf2.TextMatrix(msf2.Rows - 1, 4) = MSF1.TextMatrix(REG&, 4)
   msf2.TextMatrix(msf2.Rows - 1, 5) = MSF1.TextMatrix(REG&, 5)
   msf2.TextMatrix(msf2.Rows - 1, 6) = MSF1.TextMatrix(REG&, 6)
   If Me.Caption = "Compras" Then
      If MSF1.TextMatrix(REG&, 6) = "$" Then
           Label5 = FORMATNUM$(XVALO(Label5) + XVALO(MSF1.TextMatrix(REG&, 5)), "F12.2")
           Label8 = FORMATNUM$(XVALO(Label8) - XVALO(MSF1.TextMatrix(REG&, 5)), "F12.2")
           '
       Else
           Label13 = FORMATNUM$(XVALO(Label13) + XVALO(MSF1.TextMatrix(REG&, 5)), "F12.2")
           Label9 = FORMATNUM$(XVALO(Label9) - XVALO(MSF1.TextMatrix(REG&, 5)), "F12.2")
           
       End If
   Else
       Label5 = XVALO(Label5) + XVALO(MSF1.TextMatrix(REG&, 5))
       Label8 = XVALO(Label8) - XVALO(MSF1.TextMatrix(REG&, 5))
   End If
   '
   Call borrafila(RECALPEND09.MSF1.Row, RECALPEND09, RECALPEND09.MSF1)
   msf2.Sort = 1
End Sub
Private Sub MSF_2_DblClick()
   REG1& = MSF_2.Row
   c = MSF_2.Row
   If (MSF_2.MouseRow = 0) Then Exit Sub
'   If REG1& < 1 Then Exit Sub
   '
   'BORRO EN LA GRILLA GENERAL QUE ESTA OCULTA
   For i% = 1 To MSF1.Rows - 1
     If MSF1.TextMatrix(i%, 1) = MSF_2.TextMatrix(REG1&, 1) Then
       If MSF1.TextMatrix(i%, 2) = MSF_2.TextMatrix(REG1&, 2) Then
         If MSF1.TextMatrix(i%, 3) = MSF_2.TextMatrix(REG1&, 3) Then
           If MSF1.TextMatrix(i%, 4) = MSF_2.TextMatrix(REG1&, 4) Then
             If MSF1.TextMatrix(i%, 5) = MSF_2.TextMatrix(REG1&, 5) Then
               If MSF1.TextMatrix(i%, 6) = MSF_2.TextMatrix(REG1&, 6) Then
                 Call borrafila(i%, RECALPEND09, RECALPEND09.MSF1)
                 Exit For
               End If
             End If
           End If
         End If
       End If
     End If
   Next i%

   msf2.Rows = msf2.Rows + 1
   msf2.TextMatrix(msf2.Rows - 1, 1) = MSF_2.TextMatrix(REG1&, 1)
   msf2.TextMatrix(msf2.Rows - 1, 2) = MSF_2.TextMatrix(REG1&, 2)
   msf2.TextMatrix(msf2.Rows - 1, 3) = MSF_2.TextMatrix(REG1&, 3)
   msf2.TextMatrix(msf2.Rows - 1, 4) = MSF_2.TextMatrix(REG1&, 4)
   msf2.TextMatrix(msf2.Rows - 1, 5) = MSF_2.TextMatrix(REG1&, 5)
   msf2.TextMatrix(msf2.Rows - 1, 6) = MSF_2.TextMatrix(REG1&, 6)
   
   If Me.Caption = "Compras" Then
      If MSF_2.TextMatrix(REG1&, 6) = "$" Then
           Label5 = FORMATNUM$(XVALO(Label5) + XVALO(MSF_2.TextMatrix(REG1&, 5)), "F12.2")
           Label8 = FORMATNUM$(XVALO(Label8) - XVALO(MSF_2.TextMatrix(REG1&, 5)), "F12.2")
       Else
           Label13 = FORMATNUM$(XVALO(Label13) + XVALO(MSF_2.TextMatrix(REG1&, 5)), "F12.2")
           Label9 = FORMATNUM$(XVALO(Label9) - XVALO(MSF_2.TextMatrix(REG1&, 5)), "F12.2")
       End If
   Else
       Label5 = XVALO(Label5) + XVALO(MSF_2.TextMatrix(REG1&, 5))
       Label8 = XVALO(Label8) - XVALO(MSF_2.TextMatrix(REG1&, 5))
   End If
   
'   Label8 = XVALO(Label8) - XVALO(MSF_2.TextMatrix(REG1&, 5))
   
   

   
   
'   Label8 = XVALO(Label8) - XVALO(MSF_2.TextMatrix(REG1&, 5))
'   Label5 = XVALO(Label5) + XVALO(MSF_2.TextMatrix(REG1&, 5))
   'BORRO EN LA LISTA FILTRADA
   '
   Call borrafila(RECALPEND09.MSF_2.Row, RECALPEND09, RECALPEND09.MSF_2)

End Sub

Private Sub MSF2_DblClick()
   If (msf2.MouseRow = 0) Then Exit Sub
   
   REG& = msf2.Row
'   If REG& < 1 Then Exit Sub
   MSF1.Rows = MSF1.Rows + 1
   MSF1.TextMatrix(MSF1.Rows - 1, 1) = msf2.TextMatrix(REG&, 1)
   MSF1.TextMatrix(MSF1.Rows - 1, 2) = msf2.TextMatrix(REG&, 2)
   MSF1.TextMatrix(MSF1.Rows - 1, 3) = msf2.TextMatrix(REG&, 3)
   MSF1.TextMatrix(MSF1.Rows - 1, 4) = msf2.TextMatrix(REG&, 4)
   MSF1.TextMatrix(MSF1.Rows - 1, 5) = msf2.TextMatrix(REG&, 5)
   MSF1.TextMatrix(MSF1.Rows - 1, 6) = msf2.TextMatrix(REG&, 6)
   
   
   If Me.Caption = "Compras" Then
      If msf2.TextMatrix(REG&, 6) = "$" Then
           Label5 = FORMATNUM$(XVALO(Label5) - XVALO(msf2.TextMatrix(REG&, 5)), "F12.2")
           Label8 = FORMATNUM$(XVALO(Label8) + XVALO(msf2.TextMatrix(REG&, 5)), "F12.2")
       Else
           Label13 = FORMATNUM$(XVALO(Label13) - XVALO(msf2.TextMatrix(REG&, 5)), "F12.2")
           Label9 = FORMATNUM$(XVALO(Label9) + XVALO(msf2.TextMatrix(REG&, 5)), "F12.2")
       End If
   Else
       Label5 = XVALO(Label5) - XVALO(msf2.TextMatrix(REG&, 5))
       Label8 = XVALO(Label8) + XVALO(msf2.TextMatrix(REG&, 5))
   End If

   
'   Label8 = XVALO(Label8) + XVALO(MSF2.TextMatrix(REG&, 5))
'   Label5 = XVALO(Label5) - XVALO(MSF2.TextMatrix(REG&, 5))

   Call borrafila(RECALPEND09.msf2.Row, RECALPEND09, RECALPEND09.msf2)
   Call TXTBUSCAR_Change
End Sub
Private Sub TXTBUSCAR_Change()
    '
    If TRIM$(TXTBUSCAR) = "" Then
      MSF_2.Visible = False
      MSF1.Visible = True
      Exit Sub
    End If
    '
    MSF_2.Rows = 2
    For i& = 1 To MSF_2.Cols - 1
       MSF_2.TextMatrix(1, i&) = ""
    Next i&
    REG& = 0
    MSF_2.Visible = True
    MSF_2.ZOrder 0
    MSF_2.Cols = MSF1.Cols
    MSF_2.Redraw = False
    MSF1.Visible = False
    '
    'RECORRO EL MSF QUE TIENE TODO E IMPRIMO EN msf  LOS QUE MACHEAN
    For i& = 1 To MSF1.Rows - 1
      For J& = 1 To MSF1.Cols - 1
        If (InStr(UCase$(MSF1.TextMatrix(i&, J&)), UCase$(TXTBUSCAR))) Then
          If (TRIM$(CMBFILTRO.TEXT)) = "" Or TRIM$(UCase$(CMBFILTRO)) = TRIM$(UCase$(MSF1.TextMatrix(0, J&))) Then
            REG& = REG& + 1
            MSF_2.Rows = REG& + 1
            For h& = 1 To MSF_2.Cols - 1
              MSF_2.TextMatrix(REG&, h&) = MSF1.TextMatrix(i&, h&)
            Next h&
          End If
          If (TRIM$(CMBFILTRO.TEXT)) = "" Then Exit For
        End If
      Next J&
    Next i&
    '
    MSF_2.Redraw = True
    
    '
End Sub

Private Sub TXTBUSCAR_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      TXTBUSCAR.TEXT = ""
   End If
End Sub
