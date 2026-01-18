VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form CTRLCHEQENTRE 
   BackColor       =   &H00C7D9CC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control de Cheques Entregados"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   10620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid GRID 
      Height          =   2955
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   10395
      _ExtentX        =   18336
      _ExtentY        =   5212
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   -2147483646
      BackColorSel    =   8421504
      BackColorBkg    =   16777215
      FillStyle       =   1
      GridLines       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
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
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "CTRLCHEQENTRE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
End Sub

Private Sub GRID_DblClick()

    If GRID.Row > 0 Then
        sRecibo$ = GRID.TextMatrix(GRID.Row, 1)
        sDatosCheque$ = GRID.TextMatrix(GRID.Row, 2)
        sImporteCheque$ = GRID.TextMatrix(GRID.Row, 3)
        CondicionSql$ = "SELECT * FROM CAJABANC WHERE CODICOM_LAR='" & sRecibo$ & "' AND DESCRIMOV='" & sDatosCheque$ & "' AND VALABSCOMP=" & TRIM$(sImporteCheque$)
        
        Dim RstCheques As ADODB.Recordset
        Set RstCheques = New ADODB.Recordset
        RstCheques.Open FILTSQL$(CondicionSql$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        With RstCheques
            Do While Not .EOF
                !FEDEPOSENT = CVDAT(FECHANUM("01/01/80"))
                !Document = ""
                !DESTICHEQ = ""
                !REFESALCHE = ""
                !IDENSALCHE = ""
                !USUASALCHE = ""
                !Status = 1
                .Update
                .MoveNext
            Loop
        End With
        
        RstCheques.Close
        Set RstCheques = Nothing
        
        Call COMUNI("Estado de Cheque Recuperado")
        Unload Me
        
    End If
   
End Sub

Private Sub GRID_Click()

    Static MODO  As Boolean
    If (GRID.MouseRow = 0) Then
        ' Ordena en forma ascendente
        If MODO Then
            GRID.COL = GRID.MouseCol
            GRID.Sort = 2
            MODO = False
        ' Ordena en forma descendente
        Else
            GRID.COL = GRID.MouseCol
            GRID.Sort = 1
            MODO = True
        End If
    End If
    
End Sub
