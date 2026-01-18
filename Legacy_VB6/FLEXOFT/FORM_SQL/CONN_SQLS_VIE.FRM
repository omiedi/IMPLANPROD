VERSION 5.00
Begin VB.Form CONN_SQLS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Conexión a SQL-SERVER"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Base de Datos Inicial"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2670
      Left            =   3360
      TabIndex        =   4
      Top             =   105
      Width           =   2955
      Begin VB.ListBox LstBasedeDatos 
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
         Left            =   105
         TabIndex        =   14
         Top             =   840
         Width           =   2640
      End
      Begin VB.TextBox TxtBasedeDatos 
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
         Left            =   105
         TabIndex        =   13
         Top             =   420
         Width           =   2640
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Servidor SQL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2670
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   3165
      Begin VB.CommandButton Command3 
         Caption         =   "Comprobar Conexión"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   1680
         TabIndex        =   18
         Top             =   1995
         Width           =   1275
      End
      Begin VB.CommandButton cmdServidorSql 
         Caption         =   "Detectar Servidor SQL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   210
         TabIndex        =   17
         Top             =   1995
         Width           =   1380
      End
      Begin VB.TextBox TxtServidorSQL 
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
         Left            =   210
         TabIndex        =   12
         Top             =   420
         Width           =   2745
      End
      Begin VB.ListBox LstServidoresSQL 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Left            =   210
         TabIndex        =   11
         Top             =   840
         Width           =   2745
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00004000&
      Height          =   6030
      Left            =   6510
      ScaleHeight     =   5970
      ScaleWidth      =   2295
      TabIndex        =   7
      Top             =   0
      Width           =   2355
      Begin VB.CommandButton Command4 
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
         Left            =   120
         Picture         =   "CONN_SQLS_VIE.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Cancela la Conexión"
         Top             =   2205
         Width           =   600
      End
      Begin VB.CommandButton command1 
         Caption         =   "Exit"
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
         Left            =   120
         Picture         =   "CONN_SQLS_VIE.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Cierra - Abandona la Aplicación"
         Top             =   210
         Width           =   600
      End
      Begin VB.CommandButton Command2 
         Caption         =   "O.K."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Left            =   120
         Picture         =   "CONN_SQLS_VIE.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Comprueba la Conexión"
         Top             =   2940
         Width           =   600
      End
      Begin VB.Image Image6 
         Height          =   390
         Left            =   -315
         Picture         =   "CONN_SQLS_VIE.frx":075E
         ToolTipText     =   "Doble-Click para (Re)-Habilitar Estación de Trabajo"
         Top             =   3885
         Width           =   1515
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   330
         Left            =   90
         TabIndex        =   10
         Top             =   1520
         Width           =   645
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Modo de Autentificación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   105
      TabIndex        =   1
      Top             =   2835
      Width           =   6210
      Begin VB.TextBox TxtPWD 
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
         IMEMode         =   3  'DISABLE
         Left            =   3465
         PasswordChar    =   "*"
         TabIndex        =   16
         Top             =   990
         Width           =   2535
      End
      Begin VB.TextBox TxtUsuario 
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
         Left            =   3465
         TabIndex        =   15
         Top             =   570
         Width           =   2535
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Modo Usuario SQL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   1
         Left            =   420
         TabIndex        =   6
         Top             =   840
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Windows"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   0
         Left            =   420
         TabIndex        =   5
         Top             =   420
         Width           =   1590
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Password:"
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
         Left            =   1785
         TabIndex        =   3
         Top             =   1050
         Width           =   1590
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Username:"
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
         Left            =   1785
         TabIndex        =   2
         Top             =   630
         Width           =   1590
      End
   End
End
Attribute VB_Name = "CONN_SQLS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MiConn As ADODB.Connection
Dim mirs As ADODB.Recordset
Dim strsql As String
Dim HayError As Boolean
Dim miwindows As Boolean
Dim i As Integer
'
Function VerBases()
    LstBasedeDatos.Clear
    strsql = ("Select name from sysdatabases where sid<>0x01") ' con esta condicion no muestro las bases propias del sql
    Set mirs = MiConn.Execute(strsql)
    If Not (mirs.EOF And mirs.BOF) Then
       Do While Not mirs.EOF
          LstBasedeDatos.AddItem (mirs!Name)
          mirs.MoveNext
       Loop
    End If
End Function
  
Private Sub cmdServidorSql_Click()
    Screen.MousePointer = vbHourglass
    Dim oNames As Object
    Dim oSQLApp As Object
    Set oSQLApp = CreateObject("SQLDMO.Application")
    Set oNames = oSQLApp.ListAvailableSQLServers()  'corre la funcion que le trae los nombres de las instancias por defaul trabajo con Local
    
    LstServidoresSQL.Clear
    If oNames.Count <> 0 Then
        For i = 1 To oNames.Count
           LstServidoresSQL.AddItem oNames.Item(i)
        Next i
      Else
        MsgBox "No se Han Encontrado Instancias" & vbCrLf & "de Servidores SQL Instalados"
    End If
    
    Set oSQLApp = Nothing
    Set oNames = Nothing
    Screen.MousePointer = vbNormal
End Sub

Private Sub command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    '
    Set MiConn = New ADODB.Connection
    '
    ' validaciones **********************************************************
    If Trim$(TxtServidorSQL) = "" Then
       MsgBox "Debe Identificar Servidor SQL"
       Exit Sub
    End If
    '
    If Trim$(TxtBasedeDatos) = "" Then
       MsgBox "Debe Establecer B.Datos Inicial."
       Exit Sub
    End If
    '
    If miwindows = False Then
      If Trim$(TxtUsuario) = "" Or Trim$(TxtPWD) = "" Then
         MsgBox "Para Autenticar en Modo Usuario SQL" & vbCrLf & "Debe Establecer Username y Password."
         Exit Sub
      End If
    End If
    '
    ' fin de validaciones **************************************************
    '
    TXT$ = ""
    TXT$ = "Provider=SQLOLEDB.1;"
    '
    If miwindows = True Then
        ' autenticación windows
        TXT$ = TXT$ + "Integrated Security=SSPI;Persist Security Info=True;"
        TXT$ = TXT$ + "Data Source=" & TxtServidorSQL
      Else
        ' autenticacion sql
        TXT$ = TXT$ + "Persist Security Info=False;"
        TXT$ = TXT$ + "user id=" + TxtUsuario.Text + ";"
        TXT$ = TXT$ + "pwd=" + TxtPWD.Text + ";"
        TXT$ = TXT$ + "Data Source=" & TxtServidorSQL
    End If
    '
    TXT$ = TXT$ + ";Initial Catalog=" + TxtBasedeDatos
    TXT$ = TXT$ + ";Connect Timeout=999"
    MiConn.ConnectionString = TXT$
    '
    On Error Resume Next
    MiConn.Open
    '
    If Err Then
        On Error GoTo 0
        MsgBox "Imposible Abrir Conexion con " + TXT$, vbCritical
      Else
        On Error GoTo 0
        Call SAVEFILE(LUDAT$ + "CONNSQLS.DRV", TXT$)
        Hide
    End If
    Exit Sub
    '
End Sub

Private Sub Command3_Click()
    '
    Set MiConn = New ADODB.Connection
    Dim MiError As ADODB.Error
    On Error GoTo RutErr
    HayError = False
    '
    ' validaciones **********************************************************
    If Trim$(TxtServidorSQL) = "" Then
       MsgBox "Debe Identificar Servidor SQL"
       Exit Sub
    End If
    '
    If miwindows = False Then
      If Trim$(TxtUsuario) = "" Or Trim$(TxtPWD) = "" Then
         MsgBox "Para Autenticar en Modo Usuario SQL" & vbCrLf & "Debe Establecer Username y Password."
         Exit Sub
      End If
    End If
    '
    ' fin de validaciones **************************************************
    '
    TXT$ = ""
    TXT$ = "Provider=SQLOLEDB.1;"
    '
    If miwindows = True Then
        ' autenticación windows
        TXT$ = TXT$ + "Integrated Security=SSPI;Persist Security Info=True;"
        TXT$ = TXT$ + "Data Source=" & TxtServidorSQL
      Else
        ' autenticacion sql
        TXT$ = TXT$ + "Persist Security Info=False;"
        TXT$ = TXT$ + "user id=" + TxtUsuario.Text + ";"
        TXT$ = TXT$ + "pwd=" + TxtPWD.Text + ";"
        TXT$ = TXT$ + "Data Source=" & TxtServidorSQL
    End If
    '
    MiConn.ConnectionString = TXT$
    MiConn.CommandTimeout = 999
    MiConn.Open
    On Error GoTo 0
    '
    '
RutErr:
    For Each MiError In MiConn.Errors
       HayError = True
       MsgBox ("Error VB:" & MiError.Number & vbCrLf _
       & "Error SQL:" & MiError.SQLState & vbCrLf _
       & "Generado Por:" & MiError.Source & vbCrLf _
       & "Descripción:" & MiError.Description)
    Next
    '
    If HayError = False Then
       VerBases
       MsgBox "Conexion Verificada con " & TXT$
    End If
    
End Sub

Private Sub Command4_Click()
   Call SAVEFILE(LUDAT$ + "CONNSQLS.DRV", "")
   Unload Me
End Sub

Private Sub Form_Load()
   'Visible = True
   Caption = Caption + " - " + EMPREAC$
   Call cmdServidorSql_Click
   '
   CONNSTRI$ = Trim$(LOADFILE$(LUDAT$ + "CONNSQLS.DRV"))
   If CONNSTRI$ <> "" Then
       Set MiConn = New ADODB.Connection
       MiConn.ConnectionString = CONNSTRI$
       On Error Resume Next
       MiConn.Open
       '
       If Err Then
           On Error GoTo 0
           TxtServidorSQL = ""
           TxtBasedeDatos = ""
           MsgBox "Imposible Conectar con: " + vbCrLf + CONNSTRI$
           MiConn.Close
           Exit Sub
         Else
           On Error GoTo 0
           TXT$ = CONNSTRI$
           PINI% = InStr(TXT$, "Data Source=") + 12
           TXT1$ = Mid$(TXT$, PINI%)
           PFIN% = InStr(TXT1$, ";")
           TxtServidorSQL = Left$(TXT1$, PFIN% - 1)
           PINI% = InStr(TXT$, "Initial Catalog=") + 16
           TXT1$ = Mid$(TXT$, PINI%)
           PFIN% = InStr(TXT1$, ";")
           TxtBasedeDatos = Left$(TXT1$, PFIN% - 1)
       End If
    End If
    '
End Sub

Private Sub LstBasedeDatos_Click()
  TxtBasedeDatos.Text = LstBasedeDatos.Text
End Sub

Private Sub LstServidoresSQL_Click()
    '
    LstBasedeDatos.Clear
    TxtBasedeDatos.Text = ""
    TxtUsuario.Text = ""
    TxtPWD.Text = ""
    TxtServidorSQL.Text = LstServidoresSQL.Text
    '
End Sub

Private Sub Option1_Click(Index As Integer)
    '
    If Option1(0).Value = True Then
      TxtUsuario.Enabled = False
      TxtPWD.Enabled = False
      Label6.Enabled = False
      Label7.Enabled = False
      miwindows = True
    Else
      miwindows = False
      TxtUsuario.Enabled = True
      TxtPWD.Enabled = True
      Label6.Enabled = True
      Label7.Enabled = True
    End If
    '
End Sub
