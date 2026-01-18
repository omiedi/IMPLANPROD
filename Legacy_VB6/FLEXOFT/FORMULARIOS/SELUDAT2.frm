VERSION 5.00
Begin VB.Form SELUDAT2 
   BackColor       =   &H00FFFFE0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Opcion Multiempresa - Seleccion"
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4515
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "End"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   3600
      Picture         =   "SELUDAT2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Terminar Sesión de Trabajo - Cerrar Sistema"
      Top             =   210
      Width           =   750
   End
   Begin VB.ListBox List2 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1620
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Visible         =   0   'False
      Width           =   3165
   End
   Begin VB.Line Line2 
      X1              =   3375
      X2              =   3375
      Y1              =   210
      Y2              =   1830
   End
   Begin VB.Line Line1 
      X1              =   210
      X2              =   4515
      Y1              =   1830
      Y2              =   1830
   End
   Begin VB.Image Image3 
      Height          =   510
      Left            =   2970
      Picture         =   "SELUDAT2.frx":014A
      Top             =   1320
      Width           =   2010
   End
End
Attribute VB_Name = "SELUDAT2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Close: End
End Sub

Private Sub Form_Load()
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
End Sub

Sub List2_DblClick()
    '
    REMPRE$ = List2.Text
    EMPREAC$ = Left$(Trim$(Left$(REMPRE$, 30)) + Space$(30), 30) ' ES LA EMPRESA
    CodiEmp% = List2.ListIndex: CodiSuc% = 0
    If App.EXEName <> "FLOGON" Then
       Call validaMultiEmpXUsuario
    End If
    If InStr(1, EMPREAC$, "LAFE") > 0 Then CodiEmp% = CodiEmp% + 1
    LUSELEC$ = Trim$(Mid$(REMPRE$, 51, 46)) 'ES LA RUTA DE CARPETA DE DATOS
    '
    TCONFI$ = LOADFILE$(CurDir + "\" + "FLEXCFG.DAT")
    Call REPLA(TCONFI$, LUSELEC$ + Space$(45), 51, 45)
    Call SAVEFILE(CurDir + "\" + "FLEXCFG.DAT", TCONFI$)
    '
    RAMSTRI$ = Left$(LOADFILE(LUCOM$ + "FLEXRAM.DRV") + String$(512, 0), 512)
    Call REPLA(RAMSTRI$, Chr$(CodiEmp%), 191, 1)
    Call REPLA(RAMSTRI$, Chr$(CodiSuc%), 192, 1)
    Call REPLA(RAMSTRI$, EMPREAC$, 193, 30)
    Call SAVEFILE(LUCOM$ + "FLEXRAM.DRV", RAMSTRI$)
    Hide
    '
End Sub

Sub validaMultiEmpXUsuario()
   'VALIDA SI SE CONFIGURO CONTROL DE INGRESO DE USUARIOS AL SISTEMA PARA MULTIEMPRESA
   'ESTA VALIDACION SE CONFIGURA EN EL FLEXO09 BOTON 'M-E'
   'SI ES SUPERVISOR NO CONSIDERA LA VALIDACION
   'SI TIENE DERECHOS DEJA INGRESAR SI NO REALIZA UN CALL FINAL.
   'EL NOMBRE DEL DERECHO VARIA SEGUN LA EMPRESA EMPRES00 O EMPRES01 0 .....
'   USNU% = USNBR% Mod 100
'   If USNU% = 1 Then Exit Sub
   
   'MsgBox "1"
   SQLX$ = "SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES WITH(NOLOCK) "
   SQLX$ = SQLX$ + " WHERE TABLE_NAME = 'FLEXCRL'"
   Dim RS As New ADODB.Recordset
   On Error Resume Next
   Set RS = MiConn.Execute(SQLX$)
   If Err > 0 Then On Error GoTo 0: Exit Sub
   If Not RS.EOF Then EXISTE_LA_TABLA% = RS.Fields(0)
   RS.Close
   Set RS = Nothing
   If EXISTE_LA_TABLA% > 0 Then
        '
        COEMP$ = "00" & Trim$(Str$(CodiEmp%))
        COEMP$ = Right$(COEMP$, 2)
        '
        PROCE$ = "EMPRES" & COEMP$
        CLAVE$ = COEMP$
        
        CONDI$ = "PROBUS = '" & PROCE$ & "' AND CLABUS = '" & CLAVE$ & "'"
        SQLX$ = "SELECT VALCONTROL FROM FLEXCRL WHERE " & CONDI$
        Set RS = MiConn.Execute(SQLX$)
        If Not RS.EOF Then VALOX$ = RS!VALCONTROL
        '
        If Val(VALOX$) > 0 Then
            If DERACCE%("MUEMPU" & COEMP$, "DERECHOS DE USUARIO POR EMPRESA") < 2 Then End: Close
        End If
        RS.Close
        Set RS = Nothing
        '
  End If
     'MsgBox "3"
End Sub
Private Sub List2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
       Call List2_DblClick
     ElseIf KeyAscii = 27 Then
       Call Command1_Click
   End If
End Sub
