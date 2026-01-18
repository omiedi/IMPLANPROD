VERSION 5.00
Begin VB.Form VERPESA09 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Verificación Material en Tránsito"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   9840
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Verificados:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2715
      Left            =   150
      TabIndex        =   19
      Top             =   3435
      Width           =   8640
      Begin VB.Frame Frame8 
         Caption         =   "Frame3"
         Height          =   2200
         Left            =   6975
         TabIndex        =   25
         Top             =   360
         Width           =   15
      End
      Begin VB.Frame Frame9 
         Caption         =   "Frame3"
         Height          =   2200
         Left            =   6480
         TabIndex        =   24
         Top             =   360
         Width           =   15
      End
      Begin VB.Frame Frame7 
         Caption         =   "Frame3"
         Height          =   2200
         Left            =   5010
         TabIndex        =   23
         Top             =   345
         Width           =   15
      End
      Begin VB.Frame Frame10 
         Caption         =   "Frame3"
         Height          =   2200
         Left            =   1860
         TabIndex        =   21
         Top             =   390
         Width           =   15
      End
      Begin VB.ListBox List2 
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
         Height          =   2040
         Left            =   210
         MousePointer    =   1  'Arrow
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   20
         Top             =   560
         Width           =   8295
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"VERPESA09.frx":0000
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   210
         TabIndex        =   22
         Top             =   330
         Width           =   8295
      End
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3120
      TabIndex        =   18
      Top             =   150
      Width           =   405
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   920
      Left            =   8790
      ScaleHeight     =   915
      ScaleWidth      =   1065
      TabIndex        =   16
      Top             =   4710
      Width           =   1065
      Begin VB.CommandButton Command2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   150
         Picture         =   "VERPESA09.frx":0090
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Verificación en Central de Pesadas"
         Top             =   105
         Width           =   800
      End
   End
   Begin VB.CommandButton Command4 
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
      Height          =   300
      Left            =   3510
      TabIndex        =   15
      Top             =   135
      Width           =   175
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   0
      Top             =   0
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4200
      TabIndex        =   12
      Top             =   6270
      Width           =   4590
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Pendientes:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2715
      Left            =   135
      TabIndex        =   3
      Top             =   600
      Width           =   8640
      Begin VB.Frame Frame5 
         Caption         =   "Frame3"
         Height          =   2200
         Left            =   6960
         TabIndex        =   10
         Top             =   390
         Width           =   15
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame3"
         Height          =   2200
         Left            =   6480
         TabIndex        =   9
         Top             =   345
         Width           =   15
      End
      Begin VB.Frame Frame6 
         Caption         =   "Frame3"
         Height          =   2200
         Left            =   5025
         TabIndex        =   8
         Top             =   360
         Width           =   15
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   2200
         Left            =   1885
         TabIndex        =   7
         Top             =   360
         Width           =   15
      End
      Begin VB.ListBox List1 
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
         Height          =   2040
         Left            =   210
         MousePointer    =   1  'Arrow
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   560
         Width           =   8295
      End
      Begin VB.Label Label25 
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"VERPESA09.frx":039A
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   210
         TabIndex        =   6
         Top             =   330
         Width           =   8295
      End
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00400000&
      Height          =   6735
      Left            =   8925
      ScaleHeight     =   6675
      ScaleWidth      =   915
      TabIndex        =   0
      Top             =   -15
      Width           =   975
      Begin VB.CommandButton AGenRep 
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
         Left            =   90
         Picture         =   "VERPESA09.frx":042A
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Acceso a Consultas Pre-Configuradas"
         Top             =   3675
         Width           =   650
      End
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
         Left            =   90
         Picture         =   "VERPESA09.frx":0734
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Salir"
         Top             =   135
         Width           =   650
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   105
         Picture         =   "VERPESA09.frx":087E
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   850
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   135
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   1
         Top             =   5895
         Width           =   650
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
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "VERPESA09.frx":0B88
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   6135
         Width           =   1515
      End
   End
   Begin VB.Label Label10 
      BackColor       =   &H00E0E0E0&
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
      Height          =   285
      Left            =   3885
      TabIndex        =   14
      Top             =   150
      Width           =   2640
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Deposito (Central de Pesadas):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   -465
      TabIndex        =   13
      Top             =   195
      Width           =   3495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Verifique Código de Barras >>>>>>>>:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   11
      Top             =   6315
      Width           =   5160
   End
End
Attribute VB_Name = "VERPESA09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub AGenRep_Click()
    AGenRep.Enabled = False
    Call ACONSUL_PRE(TRIM$(UCase$(Name)))
    AGenRep.Enabled = True
End Sub
Private Sub AGenRep_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   '
   If Button = 2 Then
      Call CONECTA_CONSUL_PRE(TRIM$(UCase$(Name)))
   End If
   '
End Sub


Private Sub Command1_Click()
   '
   Call CIERRARCH("*.*")
   Unload Me
   '
End Sub

Private Sub Command14_Click()
 If DERACCE%("CFGRECE", "Configuración de Recepción") > 1 Then
    Timer1.Enabled = False
    Command14.Enabled = False
    CFGRECEP.Show 1
    Command14.Enabled = True
 End If
End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    Timer1.Enabled = False
    '
    If List1.ListCount > 0 Then
       COMUNI ("No se han Validado Todos los Bultos")
       GoTo 99
    End If
    '
    If List2.ListCount < 1 Then
       COMUNI ("Nada Para Transferir")
       GoTo 99
    End If
     
    For I = 0 To List2.ListCount - 1
       TTXX$ = List2.List(I)
       CIXI% = CODINT(Mid$(TTXX$, 1, 16))
       LOTE$ = Mid$(TTXX$, 47, 14)
       CONDI$ = "COD_INTE= " & CIXI% & " AND IDENLOTE='" & LOTE$ & "' AND DEPOMOVI = " & XVALO(Text3) & " AND CANTINGRE > 0  AND VERIFICADO = -1"
       Call ACTUREGISTRO("MOVISTO", "VERIFICADO", CONDI$, 0, REG&)
    Next I
     
    
    COMUNI ("Verificación Registrada Correctamente")
    Screen.MousePointer = 1
    Unload Me
    Exit Sub
    '
99  Command2.Enabled = True
    Screen.MousePointer = 1
    '
End Sub



Private Sub Command4_Click()
  Command4.Enabled = False
  '
  Call SELECHO("TADEPOSI")
  DEPO$ = VALACT1$("TADEPOSI")
    
  If DEPO$ <> "" Then
      Text3 = DEPO$
  End If
  '
  Command4.Enabled = True
End Sub


Private Sub Form_Activate()
  Timer1.Enabled = True
End Sub

Private Sub Form_Deactivate()
  Timer1.Enabled = False
End Sub

Private Sub Form_Load()
    '
    '\\\OT-08-1071-FL-24/11/08/// Carga la lista
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If EPAC$ <> "" Then
       Caption = Caption + "  -  " + EPAC$
    End If
    Call GRATITFOR(Caption)
    '
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
       '
    If KeyAscii > 13 Then
      Exit Sub
    End If
    '
    TENTRA$ = CONVERLECBAR(Text1)
    '
    For KKI% = 1 To List2.ListCount
       List2.Selected(KKI% - 1) = False
    Next KKI%
    '
    For KKI% = 1 To List1.ListCount
'TLIS$ = UCase(TRIM$(Left$(List1.List(KKI% - 1), 28)))
'TTZZ$ = "*"
'For KII% = 1 To Len(TLIS$): TTZZ$ = TTZZ$ + Str$(Asc(Mid$(TLIS$, KII%, 1))): Next KII%
'TTZZ$ = TTZZ$ + vbCrLf
'For KII% = 1 To Len(TENTRA$): TTZZ$ = TTZZ$ + Str$(Asc(Mid$(TENTRA$, KII%, 1))): Next KII%
'If TLIS$ = TENTRA$ Then TTZZ$ = TTZZ$ + vbCrLf + "IGUALES"
'MsgBox TTZZ$
       
       CODIGOX$ = Left$(List1.List(KKI% - 1), 16)
       LOTE$ = Mid(List1.List(KKI% - 1), 47, 12)
       CADENA$ = CODIGOX$ + TRIM(LOTE$)
       'If TLIS$ = TENTRA$ Then
       If CADENA$ = TENTRA$ Then
          TTXX$ = List1.List(KKI% - 1)
          List1.RemoveItem (KKI% - 1)
          List2.AddItem TTXX$
          List2.ListIndex = List2.ListCount - 1
          List1.ListIndex = (-1)
          GoTo 20
       End If
    Next KKI%
    '
20  For KKI% = 1 To List2.ListCount
       CODIGOX$ = Left$(List2.List(KKI% - 1), 16)
       LOTE$ = Mid(List2.List(KKI% - 1), 47, 12)
       CADENA$ = CODIGOX$ + TRIM(LOTE$)
'       If TRIM$(Left$(List2.List(KKI% - 1), 28)) = TENTRA$ Then
       If CADENA$ = TENTRA$ Then
         
          List2.ListIndex = KKI% - 1
          List2.Selected(KKI% - 1) = True
          List2.Refresh
          List1.ListIndex = (-1)
          Text1 = ""
          Text1.SetFocus
          Exit Sub
       End If
    Next KKI%
    '
    List1.ListIndex = -1
    List2.ListIndex = -1
    Call MENSERR(24, "Lote no Corresponde a Lista de Picking Activa")
    Text1 = ""
    Text1.SetFocus
    '
End Sub

Private Sub Text3_Change()
    DEPOX% = XVALO(Text3)
    DEPOT$ = ECOARCH$("TADEPOSI", Chr$(DEPOX%))
    
    If TRIM$(Text3) = "" Then
       Label10 = ""
       List1.Clear
       List2.Clear
       GoTo 99
    End If
    
    If DEPOT$ <> "" Then
       Label10 = DEPOT$
    Else
       COMUNI ("Deposito Inexistente")
       Text3 = ""
       Label10 = ""
       Text3.SetFocus
       GoTo 99
    End If
   
    '
    List1.Clear: List2.Clear: Label4 = ""
    Call APERBASDAT("STOCKS")
    Label11 = ECOARCH("TADEPOSI", Chr$(XVALO(Label7)))
    CONDI$ = ""
    
    SQLX$ = "SELECT Cod_Inte, idenLote, CantIngre "
    SQLX$ = SQLX$ + " FROM MOVISTO"
    SQLX$ = SQLX$ + " WHERE DEPOMOVI=" & XVALO(Text3) & " AND CantIngre > 0  AND VERIFICADO = -1"
    SQLX$ = SQLX$ + " ORDER BY Cod_Inte ASC"
    '
    Dim LISTCENTRAL As ADODB.Recordset
    Set LISTCENTRAL = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    JJ& = 0
    With LISTCENTRAL
       Do While Not .EOF
         CODD% = XVALO(!COD_INTE)
         KILOS = XVALO(!CantIngre)
         LOTEE$ = SAFETEXT(!idenlote)
         UME$ = UNIMED$(CODD%)
         DESCRI$ = DESCRIT0$(CODD%)
         '
         TTXX$ = AJUSTI$(CODEXT$(CODD%), 16)
         TTXX$ = TTXX$ + AJUSTI$(DESCRI$, 25)
         TTXX$ = TTXX$ + Space(5)
         TTXX$ = TTXX$ + AJUSTI$(LOTEE$, 14)
         TTXX$ = TTXX$ + AJUSTI$(UME$, 4)
         KILOS1$ = FORMATSTAN(KILOS, "F11", CODD%)
         TTXX$ = TTXX$ + KILOS1$ 'FORMATNUM$(KILOS, "F11." + CNTDC$)
         List1.AddItem TTXX$
         
       .MoveNext
       Loop
    End With
99  '
End Sub

Private Sub List1_DblClick()
    CODX$ = Left$(List1.TEXT, 28)
    Bajar_Item (CODX$)
End Sub

Private Sub List2_DblClick()
    CODX$ = Left$(List2.TEXT, 28)
    Bajar_Item (CODX$)
End Sub

Private Sub Bajar_Item(CODX$)
    If Len(CODX$) < 17 Then
       Exit Sub
    End If
    '
    TENTRA$ = TRIM$(CODX$)
    For U& = 1 To Len(TENTRA$)
      If Asc(Mid$(TENTRA$, U&, 1)) = 39 Then
        Mid$(TENTRA$, U&, 1) = Chr$(45)
      End If
    Next U&
    '
    For KKI% = 1 To List2.ListCount
       List2.Selected(KKI% - 1) = False
    Next KKI%
    '
    For KKI% = 1 To List1.ListCount
       If TRIM$(Left$(List1.List(KKI% - 1), 28)) = TENTRA$ Then
          TTXX$ = List1.List(KKI% - 1)
          List1.RemoveItem (KKI% - 1)
          List2.AddItem TTXX$
          List2.ListIndex = List2.ListCount - 1
          List1.ListIndex = (-1)
          GoTo 20
       End If
    Next KKI%
    '
20  For KKI% = 1 To List2.ListCount
       If TRIM$(Left$(List2.List(KKI% - 1), 28)) = TENTRA$ Then
          List2.ListIndex = KKI% - 1
          List2.Selected(KKI% - 1) = True
          List2.Refresh
          List1.ListIndex = (-1)
          CODX$ = ""
          
          Exit Sub
       End If
    Next KKI%
    '
    List1.ListIndex = -1
    List2.ListIndex = -1
    '
End Sub

Private Sub Text2_GotFocus()
 Timer1.Enabled = False
End Sub


Private Sub Text2_LostFocus()
  Timer1.Enabled = True
End Sub



Private Sub Timer1_Timer()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

