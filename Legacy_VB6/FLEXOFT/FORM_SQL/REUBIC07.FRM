VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form REUBIC07 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambio de Ubicación "
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   9300
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   1440
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   40
      Top             =   1080
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00CDDADA&
      Caption         =   "CANTIDAD / UBICACIÓN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3960
      TabIndex        =   34
      Top             =   3840
      Width           =   4335
      Begin VB.Label Label39 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1920
         TabIndex        =   38
         Top             =   720
         Width           =   1545
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1920
         TabIndex        =   37
         Top             =   360
         Width           =   1545
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1150
         TabIndex        =   36
         Top             =   430
         Width           =   855
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ubicación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   730
         TabIndex        =   35
         Top             =   790
         Width           =   1035
      End
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H00CDDADA&
      Caption         =   "RE-UBICACIÓN DE LOTES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   120
      TabIndex        =   19
      Top             =   5160
      Width           =   8205
      Begin VB.TextBox Text3 
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
         Height          =   285
         Left            =   3720
         TabIndex        =   28
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Nueva Ubicación"
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
         Left            =   2400
         TabIndex        =   27
         Top             =   405
         Width           =   1530
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00CDDADA&
      Caption         =   "LOTE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3960
      TabIndex        =   14
      Top             =   3000
      Width           =   4365
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1920
         TabIndex        =   33
         Top             =   240
         Width           =   1545
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Lote / Partida"
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
         Left            =   450
         TabIndex        =   32
         Top             =   315
         Width           =   1335
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00CDDADA&
      Caption         =   "FECHAS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   10
      Top             =   3000
      Width           =   3645
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Ingreso"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1005
         TabIndex        =   22
         Top             =   525
         Width           =   855
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Fabricación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   690
         TabIndex        =   21
         Top             =   1020
         Width           =   1095
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Vencimiento"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   645
         TabIndex        =   20
         Top             =   1515
         Width           =   1215
      End
      Begin VB.Label Label38 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1680
         TabIndex        =   13
         Top             =   1440
         Width           =   1275
      End
      Begin VB.Label Label37 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1680
         TabIndex        =   12
         Top             =   435
         Width           =   1275
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1680
         TabIndex        =   11
         Top             =   945
         Width           =   1275
      End
   End
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
      Height          =   285
      Left            =   3600
      TabIndex        =   9
      Top             =   315
      Width           =   175
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00CDDADA&
      Caption         =   "DETALLE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1710
      Left            =   105
      TabIndex        =   1
      Top             =   1200
      Width           =   8205
      Begin VB.Frame Frame1 
         Caption         =   "Frame3"
         Height          =   1220
         Left            =   0
         TabIndex        =   26
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame4 
         Caption         =   "Frame3"
         Height          =   1220
         Left            =   3930
         TabIndex        =   23
         Top             =   360
         Width           =   15
      End
      Begin VB.Frame Frame14 
         Caption         =   "Frame14"
         Height          =   1220
         Left            =   2880
         TabIndex        =   18
         Top             =   315
         Width           =   15
      End
      Begin VB.Frame Frame13 
         Caption         =   "Frame3"
         Height          =   1170
         Left            =   6900
         TabIndex        =   17
         Top             =   360
         Width           =   15
      End
      Begin VB.Frame Frame12 
         Caption         =   "Frame3"
         Height          =   1170
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame11 
         Caption         =   "Frame3"
         Height          =   1170
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   1220
         Left            =   5760
         TabIndex        =   4
         Top             =   315
         Width           =   15
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   930
         Left            =   600
         MousePointer    =   1  'Arrow
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   600
         Width           =   6945
      End
      Begin VB.Label Label25 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Lote                                           F.Venc.          Ubicación                           Cantidad"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   600
         TabIndex        =   3
         Top             =   330
         Width           =   6945
      End
   End
   Begin VB.PictureBox Picture14 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   6570
      Left            =   8400
      ScaleHeight     =   6540
      ScaleWidth      =   1350
      TabIndex        =   0
      Top             =   0
      Width           =   1380
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   120
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   42
         Top             =   5760
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   43
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "O.K."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Left            =   120
         Picture         =   "REUBIC07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Graba - Aprueba la Transacción"
         Top             =   3960
         Width           =   645
      End
      Begin VB.CommandButton Command11 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "REUBIC07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Salir"
         Top             =   120
         Width           =   645
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "REUBIC07.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   870
         Width           =   645
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Kdx"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "REUBIC07.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Ficha Tipo Kardex de Movimientos Históricos"
         Top             =   1605
         Width           =   645
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   4200
      OleObjectBlob   =   "REUBIC07.frx":0A68
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "REUBIC07.frx":0C9C
      TabIndex        =   39
      Top             =   0
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5450
      TabIndex        =   25
      Top             =   385
      Width           =   1545
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   7130
      TabIndex        =   24
      Top             =   315
      Width           =   1020
   End
   Begin VB.Label Label31 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   210
      TabIndex        =   8
      Top             =   810
      Width           =   1125
   End
   Begin VB.Label Label30 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1470
      TabIndex        =   7
      Top             =   735
      Width           =   6680
   End
   Begin VB.Label Label28 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1470
      TabIndex        =   6
      Top             =   315
      Width           =   2120
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   -210
      TabIndex        =   5
      Top             =   385
      Width           =   1545
   End
End
Attribute VB_Name = "REUBIC07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   '
   Call CIERRARCH("*.*")
   Hide
   '
End Sub

Private Sub Command10_Click()
 '
   Command10.Enabled = False
   CUDI$ = UCase$(CurDir)
   PPXX% = InStr(CUDI$, "\WINDOWS")
   If PPXX% > 0 Then
      RUTAX$ = Left$(CUDI$, PPXX%) + "MANUALES\"
      If EXISTE%(RUTAX$ + "TRAZAB04.PDF") > 0 Then
         FORHLPDF.Pdf1.setShowToolbar (True)
         On Error Resume Next
         FORHLPDF.Pdf1.LOADFILE RUTAX$ + "TRAZAB04.pdf"
         If Err Then GoTo 20
         On Error GoTo 0
         FORHLPDF.Show 1
         GoTo 99
      End If
   End If
   '
20 Call MENSERR(24, "Imposible Abrir Archivo de Ayuda")
   '
99 Command10.Enabled = True
End Sub

Private Sub Command11_Click()
   '
   Call CIERRARCH("*.*")
   Unload Me
   '
End Sub

Private Sub Command2_Click()

   Command2.Enabled = False
   '
   'VALIDACIONES
   FECHA% = FECHANUM(Label13)
   FEFA% = FECHANUM(Label36)
   FEVE% = FECHANUM(Label38)
   CIXI% = CODINT%(Label28)
   IDLOTX$ = TRIM$(Label2)
   '
   If TRIM$(Label39.Caption) = "" Then
      Call MENSERR(24, "Lote Sin Ubicación Original.")
      GoTo 99
   End If
   '
   If CIXI% < 1 Then
      Call MENSERR(24, "Codigo no Válido.")
      GoTo 99
   End If
   '
   UBICAIN$ = TRIM$(Left$(UCase(Text3), 32))
   If UBICAIN$ = "" Then
    Call MENSERR(24, "Falta Ingreso de Ubicación")
    GoTo 99
   End If
      '
   CODIPROI% = XVALO(VALOBADA("LOTINGRE", "Cod_Inte", "UbicFis = '" & UBICAIN$ & "' AND CANALTA > (CANCONSU + 0.05) ORDER BY FECHALTA ASC "))
   If CODIPROI% > 0 Then
      CODEPRO$ = CODEXT$(CODIPROI%)
      OPX% = COMALTER%("La Ubicacion " + UBICAIN$ + " se Encuentra Ocupada\por Código '" + CODEPRO$ + "'.\  \¿Confirma la Posicion Elegida?\\No, Cambiar\Si, Continuar")
      '
      If OPX% = 1 Then
        GoTo 99
      End If
   End If
   '
   SQLX$ = "Select UBICFIS from LOTINGRE where IdenLote = '" & IDLOTX$ & "'" 'Selecciona el registro correspondiente al LOTE
   Dim LOTINGTMP As ADODB.Recordset
   Set LOTINGTMP = New ADODB.Recordset
   LOTINGTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   '
   With LOTINGTMP
      !UBICFIS = UBICAIN$
      .Update
   End With
   Set LOTINGTMP = Nothing
   '
   Condi$ = "IDENLOTE='" & TRIM$(IDLOTX$) & "'"
   DPX% = XVALO(VALOBADA("LOTINGRE", "DepoUbic", Condi$))
   If DPX% > 0 Then Call PRINSTIKAPRO(CIXI%, IDLOTX$, DPX%)

   Call COMUNI("Procesado")
   '
   Call LIMPLAB
   Call LIMPTEXT
   Call Label28_Change
 
99 Screen.MousePointer = 1
   Command2.Enabled = True
   '
End Sub

Private Sub Command5_Click()
    Call SELECHO("MASTER")
    If VALACT1$("MASTER") <> "" Then
       Label28 = VALACT1$("MASTER")
    End If
End Sub

Private Sub Command7_Click()
   Command7.Enabled = False
   FIKARDEX09.Show 1
   Command7.Enabled = True
End Sub

Private Sub Command9_Click()
    Call CommSel("TADEPOSI", Command9, Text6)
End Sub

Private Sub Form_Load()
    '
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
    HOII% = HOY(HOS$)
    Label13 = HOS$
    '
    If Left$(APLINVO$, 7) = "BORECEP" Then ' Or INTERPRE% = 1 Then
       Text1 = Mid$(APLINVO$, 8)
       Openforms = DoEvents
    End If
    '
End Sub

'
Private Sub Form_Unload(Cancel As Integer)
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Label28_Change()
    CIXI% = CODINT%(Label28)
    Label30 = DESCRIT0$(CIXI%)
    Call LIMPLAB
    Call LIMPTEXT
    Call ARMALILOTES(CIXI%)
End Sub

Sub ARMALILOTES(CIXI%)
    '
     FLEXCONN.Execute (FILTSQL$("SET DATEFORMAT dmy"))
    '
    Call APERBASDAT("STOCKS")
    SQLX$ = " SELECT * FROM LOTINGRE WITH(NOLOCK)"
'    SQLX$ = SQLX$ + " INNER JOIN BOLRECEP"
'    SQLX$ = SQLX$ + " ON LOTINGRE.IDENLOTE = BOLRECEP.IDENLOTE"
    SQLX$ = SQLX$ + " WHERE LOTINGRE.Cod_Inte = " & CIXI%
    SQLX$ = SQLX$ + " AND (CanAlta - CanConsu) > 0"    ' no trae los consumidos
    SQLX$ = SQLX$ + " ORDER BY FechAlta ASC, LOTINGRE.IdenLote ASC "
    '
    Dim LOTINGTEMP As ADODB.Recordset
    Set LOTINGTEMP = New ADODB.Recordset
    Set LOTINGTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    REBLA$ = Space$(256)
    JJ& = 0: List1.Clear
    With LOTINGTEMP
25     On Error Resume Next
       .MoveFirst
       If Err Then
         On Error GoTo 0
         Call MENSERR(24, "No Figuran Lotes Disponibles\para el Material Elegido.")
         GoTo 99
       End If
       '
       On Error GoTo 0
       Do While Not .EOF
         IDLOT$ = TRIM$(!IDENLOTE)
         '
         TTXX$ = REBLA$
         Call REPLA(TTXX$, IDLOT$, 1, 20) 'lote
         FEX = CVINT(!FechVenc)
         Call REPLA(TTXX$, FECHATEX$(FEX), 23, 8)                          ' Fecha vencimiento
         Call REPLA(TTXX$, SAFETEXT(!UBICFIS), 33, 16)                      ' Ubicacion
         Call REPLA(TTXX$, FORMATNUM$(XVALO(!CanAlta - !CanConsu), "F12.1"), 48, 12)  ' cantidad
         Call REPLA(TTXX$, XVALO(!DepoUbic), 70, 2)                         ' Deposito
         Call REPLA(TTXX$, SAFETEXT(!CodEnvase), 72, 25)                    ' Envase
         FEXALTA = CVINT(!FECHALTA)
         Call REPLA(TTXX$, FECHATEX$(FEXALTA), 97, 8)                       ' fecha Ingreso
         FEXFAB = CVINT(!FechFab)
         Call REPLA(TTXX$, FECHATEX$(FEXFAB), 105, 8)                       ' Fecha fabricación
            COEFI = 1: If !CanAlta > 0 Then COEFI = !Cantidad / !CanAlta
         Call REPLA(TTXX$, FORMATNUM(XVALO(COEFI * (!CanAlta - !CanConsu)), "F12.1"), 113, 12) ' Cantidad
         '
         JJ& = JJ& + 1
         List1.AddItem TTXX$
       .MoveNext
       Loop
    End With
    '
99  LOTINGTEMP.Close
    Set LOTINGTEMP = Nothing
    '
End Sub

Private Sub List1_Click()
    
    Call LIMPTEXT
    '
    TTXX$ = List1.TEXT
    '
    NCR$ = Mid(TTXX$, 1, 20)
    COD$ = Label28
    CI% = CODINT%(COD$)
    Label2.Caption = NCR$
    Label37.Caption = TRIM$(Mid$(TTXX$, 97, 8))   ' F. Ingreso
    Label36.Caption = TRIM$(Mid$(TTXX$, 105, 8))  ' F. Fabricación
    Label38.Caption = TRIM$(Mid$(TTXX$, 23, 8))   ' F. Vencimiento
    Label41.Caption = TRIM$(Mid$(TTXX$, 49, 12))  ' cantidad
    Label39.Caption = TRIM$(Mid$(TTXX$, 33, 16))  ' Ubicación
    Label34 = Label39.Caption
    '
    NCR$ = TRIM$(NCR$)
    PPXX% = InStr(NCR$, "/")
    If PPXX% > 0 Then
       NCR$ = Left$(NCR$, PPXX% - 1) ' VERIFICA LA EXISTENCIA DE LA BARRA IDENTIFICADORA DE SUBLOTE
    End If
    '
   
End Sub

Sub LIMPLAB()
    Label2.Caption = ""
    Label36.Caption = ""
    Label37.Caption = ""
    Label38.Caption = ""
    Label39.Caption = ""
    Label41.Caption = ""
End Sub

Sub LIMPTEXT()
    Text1 = ""
    Text2 = ""
    Text3 = ""
    Text4 = ""
    Text5 = ""
    Text6 = ""
    Text8 = ""
    Text9 = ""
    Text14 = ""
    Text11 = ""
    Text16 = ""
    Text12 = ""
    Text17 = ""
    Call BLANARCH("!CARFRABO")
End Sub

Private Sub Text1_Change()
'
Label32.Caption = FORMATNUM$(XVALO(Label41.Caption) - XVALO(Text1.TEXT), "F12.1")  'RESTA AL LOTE ORIGINALLOS cantidad DEL NUEVO LOTE
If XVALO(Label32.Caption) < 0 Then
    Call MENSERR(24, "Cantidad  Supera la del Lote Original !!!")
End If
'
End Sub

Private Sub Text1_LostFocus()
'
Text1.TEXT = FORMATNUM$(Text1.TEXT, "F12.1") ' FORMATEA EL TEXT1 "cantidad"
'
End Sub

Private Sub Text2_Change()
'
If XVALO(Label33.Caption) < 0 Then
    Call MENSERR(24, "Cantidad Supera la del Lote Original")
End If
'
End Sub

Private Sub Text2_LostFocus()
'
Text2.TEXT = FORMATNUM$(Text2.TEXT, "F12.1") ' FORMATEA EL TEXT2 "CANTIDAD"
'
End Sub




