VERSION 5.00
Begin VB.Form CONN_SQLS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Conexión a SQL-SERVER"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame5 
      Caption         =   "Cadena de Conexion"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   105
      TabIndex        =   24
      Top             =   5040
      Width           =   6280
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   210
         MultiLine       =   -1  'True
         TabIndex        =   25
         Top             =   315
         Width           =   5895
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Ruta / Carpeta de Datos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   105
      TabIndex        =   22
      Top             =   105
      Width           =   6280
      Begin VB.Label Label1 
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
         Left            =   630
         TabIndex        =   23
         Top             =   315
         Width           =   5475
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Base de Datos Inicial"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2670
      Left            =   3360
      TabIndex        =   4
      Top             =   945
      Width           =   3020
      Begin VB.CommandButton CmdCrearUsuario 
         Caption         =   "Crea Usuario FLEXOFT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   105
         TabIndex        =   21
         Top             =   1995
         Width           =   1275
      End
      Begin VB.CommandButton CmdCrearBase 
         Caption         =   "Crea Base de Datos SQL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   1470
         TabIndex        =   20
         Top             =   1995
         Width           =   1380
      End
      Begin VB.ListBox LstBasedeDatos 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   105
         TabIndex        =   14
         Top             =   840
         Width           =   2745
      End
      Begin VB.TextBox TxtBasedeDatos 
         BeginProperty Font 
            Name            =   "Arial"
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
         Width           =   2745
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Servidor SQL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2670
      Left            =   105
      TabIndex        =   0
      Top             =   945
      Width           =   3165
      Begin VB.CommandButton Command3 
         Caption         =   "Comprobar Conexión"
         BeginProperty Font 
            Name            =   "Arial"
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
            Name            =   "Arial"
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
            Name            =   "Arial"
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
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   210
         TabIndex        =   11
         Top             =   840
         Width           =   2745
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00004000&
      Height          =   6500
      Left            =   6510
      ScaleHeight     =   6435
      ScaleWidth      =   2295
      TabIndex        =   7
      Top             =   0
      Width           =   2355
      Begin VB.CommandButton Command4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   105
         Picture         =   "CONN_SQLS.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Cancela la Conexión a SQL SERVER"
         Top             =   4620
         Width           =   600
      End
      Begin VB.CommandButton command1 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   120
         Picture         =   "CONN_SQLS.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Cierra - Abandona la Aplicación"
         Top             =   180
         Width           =   600
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
         Picture         =   "CONN_SQLS.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Comprueba y Registra la Conexión Configurada"
         Top             =   5355
         Width           =   600
      End
      Begin VB.Image Image6 
         Height          =   390
         Left            =   -315
         Picture         =   "CONN_SQLS.frx":075E
         ToolTipText     =   "Doble-Click para (Re)-Habilitar Estación de Trabajo"
         Top             =   3990
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
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   105
      TabIndex        =   1
      Top             =   3675
      Width           =   6280
      Begin VB.TextBox TxtPWD 
         BeginProperty Font 
            Name            =   "Arial"
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
         Top             =   780
         Width           =   2640
      End
      Begin VB.TextBox TxtUsuario 
         BeginProperty Font 
            Name            =   "Arial"
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
         Top             =   360
         Width           =   2640
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Modo Usuario SQL"
         BeginProperty Font 
            Name            =   "Arial"
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
         Top             =   735
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Windows"
         BeginProperty Font 
            Name            =   "Arial"
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
         Top             =   315
         Width           =   1590
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Password:"
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
         Left            =   1785
         TabIndex        =   3
         Top             =   840
         Width           =   1590
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Username:"
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
         Left            =   1785
         TabIndex        =   2
         Top             =   420
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
Dim MIRS As ADODB.Recordset
Dim strsql As String
Dim HayError As Boolean
Dim miwindows As Boolean
Dim i As Integer
'
 
Private Sub CmdCrearBase_Click()
    '
    CREABASE.Show 1
    NOMEBASE$ = TRIM$(CREABASE.Text1.TEXT)
    If NOMEBASE$ = "" Then Exit Sub           ' CANCELADO
    '
    RUTBASE$ = TRIM$(CREABASE.Label3.Caption)
    If RUTBASE$ = "" Then Exit Sub
    If EXISTE%(RUTBASE$ + "\.") < 1 Then
       MsgBox "Carpeta '" + RUBASE$ + "' Inexistente", vbExclamation
       Exit Sub
    End If
    '
    On Error GoTo RutErr
    If MiConn.State < 1 Then
       MsgBox "Debe Primero Comprobar la Conexión."
       Exit Sub
    End If
    '
    ' crea la base en el directorio ludat$
    ' del SQLServer con los parametros por default
     strsql = "Select name from sysobjects where name = '" & NOMEBASE$ & "'"
     Set MIRS = MiConn.Execute(strsql)
     If Not (MIRS.EOF And MIRS.BOF) Then
          MsgBox "La Base de Datos '" + NOMEBASE$ + "' ya Existe."
        Else
          strsql = "CREATE DATABASE " & NOMEBASE$
          strsql = strsql & " ON (NAME = " & NOMEBASE$ & ","
          strsql = strsql & " FILENAME = '" & RUTBASE$ & "\" & NOMEBASE$ & "_Dat.mdf')"
          strsql = strsql & " LOG ON (NAME = '" & NOMEBASE$ & "_Log',"
          strsql = strsql & " FILENAME = '" & RUTBASE$ & "\" & NOMEBASE$ & "_log.ldf ')"
          MiConn.Execute (strsql)
          MsgBox "Se Ha Creado la Base de Datos"
     End If
     '
     VerBases
     '
90   On Error GoTo 0
     Exit Sub
     '
RutErr:
     For Each MiError In MiConn.Errors
        HayError = True
        MsgBox ("Fallo la Creación de la Base de Datos " & vbCrLf & vbCrLf _
        & "Error VB: " & MiError.Number & vbCrLf _
        & "Error SQL: " & MiError.SQLState & vbCrLf _
        & "Generado Por: " & MiError.Source & vbCrLf _
        & "Descripción: " & MiError.Description & vbCrLf & vbCrLf _
        & "Recuerde: " & vbCrLf _
        & "1.- Solo un usuario con nivel de Administrador puede crear una Base de Datos. " & vbCrLf _
        & "2.- El proceso de creacion de una Base de Datos debe ejecutarse en el equipo servidor SQL. " & vbCrLf _
        & "3.- La Ruta debe referenciar una carpeta valida en el equipo donde reside el motor SQL SERVER." & vbCrLf _
        & "      No son Validas unidades de red mapeadas")
     Next
     '
End Sub

Private Sub CmdCrearUsuario_Click()
   On Error GoTo RutErr
   If MiConn.State < 1 Then
      MsgBox "Debe Primero Comprobar la Conexión."
      GoTo 90
   End If
   '
   strsql = "Select * from master.dbo.syslogins where loginname = 'FLEXOFT'"
   Set MIRS = MiConn.Execute(strsql)
   If Not (MIRS.EOF And MIRS.BOF) Then
         MsgBox "El Usuario de SQL-Server 'FLEXOFT' ya Existe."
      Else
         'crea al usuario
          strsql = ("sp_addlogin @loginame='FLEXOFT'")
          strsql = strsql + (", @passwd='bauness'")
          MiConn.Execute (strsql)
          '
          'genera el rol del usuario en nuestro caso system administrator
          strsql = ("select name from syslogins where name='FLEXOFT'")
          Set MIRS = MiConn.Execute(strsql)
          If Not (MIRS.EOF And Not MIRS.BOF) Then
             strsql = ("sp_addsrvrolemember 'FLEXOFT', 'sysadmin'")
             MiConn.Execute (strsql)
          End If
          '
          MsgBox "Se ha Creado el Usuario 'FLEXOFT' en el Servidor SQL-Server con Privilegios de Administrador."
   End If
90 On Error GoTo 0
   Exit Sub
   '
RutErr:
   For Each MiError In MiConn.Errors
      HayError = True
      MsgBox ("Falló la Creación del Usuario FLEXOFT" & vbCrLf _
      & "Error VB: " & MiError.Number & vbCrLf _
      & "Error SQL: " & MiError.SQLState & vbCrLf _
      & "Generado Por: " & MiError.Source & vbCrLf _
      & "Descripción: " & MiError.Description)
   Next
   '
End Sub
Private Sub cmdServidorSql_Click()
    'SQLDMO se utiliza para traer los servidores SQL instalados en la red
    
    Screen.MousePointer = vbHourglass
    Dim oNames As Object
    Dim oSQLApp As SQLDMO.Application
    On Error Resume Next
    'Set oSQLApp = CreateObject("SQLDMO.Application")
    Set oSQLApp = New SQLDMO.Application
    If Err Then
       TTXX$ = "Error " & Err.Number & " " & Err.Description & "." & vbCrLf & vbCrLf & "En esta Estación de Trabajo no es Posible Establecer los Parámetros de Conexión en Forma AUTOMÁTICA. Establezca los Parámetros en forma Manual o Acceda desde otro Equipo de la Red LAN."
       On Error GoTo 0
       MsgBox TTXX$
       GoTo 80
    End If
    '
    Set oNames = oSQLApp.ListAvailableSQLServers()  'corre la funcion que le trae los nombres de los servidores SQL
    LstServidoresSQL.Clear
    If oNames.Count <> 0 Then
        For i = 1 To oNames.Count
           LstServidoresSQL.AddItem oNames.Item(i)
        Next i
      Else
        MsgBox "No se Han Encontrado Instancias" & vbCrLf & "de Servidores SQL Instalados"
    End If
    '
80  On Error Resume Next
    Set oSQLApp = Nothing
    Set oNames = Nothing
    Screen.MousePointer = vbNormal
    On Error GoTo 0
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub


Private Sub Command2_Click() 'OK
    '
    Command2.Enabled = False
    Set MiConn = New ADODB.Connection
    '
'    ' validaciones **********************************************************
'    If Trim$(TxtServidorSQL) = "" Then
'       MsgBox "Debe Identificar Servidor SQL"
'       Exit Sub
'    End If
'    '
'    If Trim$(TxtBasedeDatos) = "" Then
'       MsgBox "Debe Establecer B.Datos Inicial."
'       Exit Sub
'    End If
'    '
'    If miwindows = False Then
'      If Trim$(TxtUsuario) = "" Or Trim$(TxtPWD) = "" Then
'         MsgBox "Para Autenticar en Modo Usuario SQL" & vbCrLf & "Debe Establecer Username y Password."
'         Exit Sub
'      End If
'    End If
'    '
'    ' fin de validaciones **************************************************
'    '
'    TXT$ = ""
'    TXT$ = "Provider=SQLOLEDB.1;"
'    '
'    If miwindows = True Then
'        ' autenticación windows
'        TXT$ = TXT$ + "Integrated Security=SSPI;Persist Security Info=True;"
'        TXT$ = TXT$ + "Data Source=" & TxtServidorSQL
'      Else
'        ' autenticacion sql
'        TXT$ = TXT$ + "Persist Security Info=False;"
'        TXT$ = TXT$ + "user id=" + TxtUsuario.Text + ";"
'        TXT$ = TXT$ + "pwd=" + TxtPWD.Text + ";"
'        TXT$ = TXT$ + "Data Source=" & TxtServidorSQL
'    End If
'    '
'    TXT$ = TXT$ + ";Initial Catalog=" + TxtBasedeDatos
'    TXT$ = TXT$ + ";Connect Timeout=999"
'    MiConn.ConnectionString = TXT$
    Screen.MousePointer = 11
    TXT$ = REPLACAR$(Text1, "999", "9  ")
    MiConn.ConnectionString = TXT$
    '
    On Error Resume Next
    MiConn.Open
    '
    If Err Then
        DERRO$ = Err.Description
        On Error GoTo 0
        Screen.MousePointer = 1
        MsgBox "Imposible Abrir Conexion con la Siguiente Cadena de Conexion:" + vbCrLf + vbCrLf + "'" + TXT$ + "'" + vbCrLf + vbCrLf + vbCrLf + "Error Reportado:" + vbCrLf + vbCrLf + DERRO$, vbCritical
      Else
        On Error GoTo 0
        Screen.MousePointer = 1
        Call SAVEFILE(LUDAT$ + "CONNSQLS.DRV", TRAPASS2$(TXT$))
        MsgBox "Conexion Verificada OK y Registrada."
        Command2.Enabled = True
        Hide
    End If
    Command2.Enabled = True
    Exit Sub
    '
End Sub

Sub ARMACONEXION()
    '
    Set MiConn = New ADODB.Connection
    '
    ' validaciones **********************************************************
'    If TRIM$(TxtServidorSQL) = "" Then
'       MsgBox "Debe Identificar Servidor SQL"
'       Exit Sub
'    End If
'    '
'    If TRIM$(TxtBasedeDatos) = "" Then
'       MsgBox "Debe Establecer B.Datos Inicial."
'       Exit Sub
'    End If
'    '
'    If miwindows = False Then
'      If TRIM$(TxtUsuario) = "" Or TRIM$(TxtPWD) = "" Then
'         MsgBox "Para Autenticar en Modo Usuario SQL" & vbCrLf & "Debe Establecer Username y Password."
'         Exit Sub
'      End If
'    End If
'    '
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
        TXT$ = TXT$ + "user id=" + TxtUsuario.TEXT + ";"
        TXT$ = TXT$ + "pwd=" + TxtPWD.TEXT + ";"
        TXT$ = TXT$ + "Data Source=" & TxtServidorSQL
    End If
    '
    TXT$ = TXT$ + ";Initial Catalog=" + TxtBasedeDatos
    TXT$ = TXT$ + ";Connect Timeout=999"
    Text1 = TXT$
    '
End Sub

Private Sub Command3_Click() 'Comprobar Conexion
    '
    Screen.MousePointer = 11
    Set MiConn = New ADODB.Connection
    Dim MiError As ADODB.Error
    On Error GoTo RutErr
    HayError = False
    '
    ' validaciones **********************************************************
    If TRIM$(TxtServidorSQL) = "" Then
       MsgBox "Debe Identificar Servidor SQL"
       Screen.MousePointer = 1
       Exit Sub
    End If
    '
    If miwindows = False Then
      If TRIM$(TxtUsuario) = "" Or TRIM$(TxtPWD) = "" Then
         MsgBox "Para Autenticar en Modo Usuario SQL" & vbCrLf & "Debe Establecer Username y Password."
         Screen.MousePointer = 1
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
        TXT$ = TXT$ + "user id=" + TxtUsuario.TEXT + ";"
        TXT$ = TXT$ + "pwd=" + TxtPWD.TEXT + ";"
        TXT$ = TXT$ + "Data Source=" & TxtServidorSQL
    End If
    '
    MiConn.ConnectionString = TXT$
    MiConn.CommandTimeout = 999
    On Error Resume Next
    MiConn.Open
    If Err Then
       Screen.MousePointer = 1
       MsgBox "Falló la Verificacion de la Conexión."
       GoTo RutErr
    End If
    On Error GoTo 0
    '
RutErr:
    For Each MiError In MiConn.Errors
       HayError = True
       MsgBox ("Error VB: " & MiError.Number & vbCrLf _
       & "Error SQL: " & MiError.SQLState & vbCrLf _
       & "Generado Por: " & MiError.Source & vbCrLf _
       & "Descripción: " & MiError.Description)
    Next
    '
    If HayError = False Then
       VerBases
       Screen.MousePointer = 1
       MsgBox "Conexion Verificada con:" + vbCrLf + vbCrLf & TXT$
    End If
    
End Sub

Function VerBases()
    LstBasedeDatos.Clear
    strsql = ("Select name from master.dbo.sysdatabases") ' con esta condicion no muestro las bases propias del sql
    Set MIRS = MiConn.Execute(strsql)
    If Not (MIRS.EOF And MIRS.BOF) Then
       Do While Not MIRS.EOF
          LstBasedeDatos.AddItem (MIRS!Name)
          MIRS.MoveNext
       Loop
    Else
      CmdCrearUsuario_Click
    End If
End Function
Private Sub Command4_Click()
   Call SAVEFILE(LUDAT$ + "CONNSQLS.DRV", "")
   Unload Me
End Sub

Private Sub Form_Load()
   'Visible = True
   UTILIZA_SKIN% = 1
   Call APLICACIONSKINS(Me)
   '
   Caption = Caption + " - " + EMPREAC$
   Label1 = LUDAT
   '
   If FALLOCONEXION% = 1 Then Exit Sub   ' PARA QUE NO VUELVA A VERIFICAR
   FALLOCONEXION% = 0
   '
   Screen.MousePointer = 11
   Call cmdServidorSql_Click
   CONNSTRI$ = TRAPASS1$(TRIM$(LOADFILE$(LUDAT$ + "CONNSQLS.DRV")))
   '
   If CONNSTRI$ <> "" Then
       CONNSTRI1$ = CONNSTRI$
       PPXX% = InStr(CONNSTRI1$, "Timeout=999")
       If PPXX% > 1 Then
          Mid$(CONNSTRI1$, PPXX%, 11) = "Timeout=9  "
       End If
       '
       Set MiConn = New ADODB.Connection
       MiConn.ConnectionString = CONNSTRI1$
       On Error Resume Next
       MiConn.Open
       '
       Screen.MousePointer = 1
       If Err Then
           On Error GoTo 0
           TxtServidorSQL = ""
           TxtBasedeDatos = ""
           MsgBox "Imposible Conectar con: " + CONNSTRI1$
           On Error Resume Next
           MiConn.Close
           On Error GoTo 0
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
           PINI% = InStr(TXT$, "user id=") + 8
           TXT1$ = Mid$(TXT$, PINI%)
           PFIN% = InStr(TXT1$, ";")
           TxtUsuario = Left$(TXT1$, PFIN% - 1)
           If TRIM$(TxtUsuario) <> "" Then Option1(1) = True
           PINI% = InStr(TXT$, "pwd=") + 4
           TXT1$ = Mid$(TXT$, PINI%)
           PFIN% = InStr(TXT1$, ";")
           TxtPWD = Left$(TXT1$, PFIN% - 1)
       End If
       Text1 = CONNSTRI$
   End If
   '
End Sub

Private Sub LstBasedeDatos_Click()
  TxtBasedeDatos.TEXT = LstBasedeDatos.TEXT
End Sub

Private Sub LstServidoresSQL_Click()
    '
    LstBasedeDatos.Clear
    TxtBasedeDatos.TEXT = ""
    TxtUsuario.TEXT = ""
    TxtPWD.TEXT = ""
    TxtServidorSQL.TEXT = LstServidoresSQL.TEXT
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
    Call ARMACONEXION
    '
End Sub

Private Sub TxtBasedeDatos_Change()
   Call ARMACONEXION
End Sub

Private Sub TxtPWD_Change()
   Call ARMACONEXION
End Sub

Private Sub TxtServidorSQL_Change()
    Call ARMACONEXION
End Sub

Private Sub TxtUsuario_Change()
    On Error Resume Next
    MiConn.Close    ' si cambia el usuario, cierra la conexion
    On Error GoTo 0
    Call ARMACONEXION
End Sub
