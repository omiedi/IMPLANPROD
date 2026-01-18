VERSION 5.00
Begin VB.Form SINSUCBA07 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sincronización de Sucursales"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   7545
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   105
      TabIndex        =   18
      Top             =   3120
      Width           =   6375
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H80000008&
      Height          =   6660
      Left            =   6600
      ScaleHeight     =   6630
      ScaleWidth      =   1005
      TabIndex        =   14
      Top             =   0
      Width           =   1035
      Begin VB.CommandButton Command6 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   120
         Picture         =   "SINSUCBA07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Registra y Guarda la Configuración"
         Top             =   960
         Width           =   700
      End
      Begin VB.CommandButton Command7 
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
         Height          =   700
         Left            =   120
         Picture         =   "SINSUCBA07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Salir"
         Top             =   240
         Width           =   700
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Restaura Backup"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   10
      Top             =   5955
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Toma Bkp de Dropbox"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   9
      Top             =   5400
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Mueve Bkp a Dropbox"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   5955
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Genera Backup"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   5400
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   "CONFIGURACIONES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   105
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   30000
         Left            =   3960
         Top             =   1440
      End
      Begin VB.Timer Timer2 
         Interval        =   5000
         Left            =   4440
         Top             =   1440
      End
      Begin VB.TextBox Text5 
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
         Left            =   1440
         TabIndex        =   20
         Top             =   1800
         Width           =   4695
      End
      Begin VB.TextBox Text4 
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
         Left            =   1440
         TabIndex        =   16
         Top             =   720
         Width           =   4695
      End
      Begin VB.Frame Frame2 
         Caption         =   "ORIGEN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   11
         Top             =   2250
         Width           =   5415
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            Caption         =   "Central"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   960
            TabIndex        =   13
            Top             =   200
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton Option2 
            Appearance      =   0  'Flat
            Caption         =   "Sucursal"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3840
            TabIndex        =   12
            Top             =   200
            Width           =   1335
         End
      End
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
         Left            =   1440
         TabIndex        =   6
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox Text2 
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
         Left            =   1440
         TabIndex        =   4
         Top             =   1080
         Width           =   4695
      End
      Begin VB.TextBox Text1 
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
         Left            =   1440
         TabIndex        =   2
         Top             =   360
         Width           =   4695
      End
      Begin VB.Label Label5 
         Caption         =   "Acrchivo Transferencia"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   1720
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Carpeta Backup"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   780
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Horario"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   1485
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Carpeta Intercambio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   1020
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Base de Datos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   435
         Width           =   1695
      End
   End
End
Attribute VB_Name = "SINSUCBA07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim UltimoBackup As String
Dim Reintenta As Boolean
Dim BASEDA As String
Dim RUTADB As String
Dim RUTABK As String
Dim HORARIO As String
Dim RUTASIS As String
Dim ORIGEN  As String
Dim SUCURSAL As String
Dim CarpetaDefault As String
Dim FLEXCONN As ADODB.Connection

Private Sub Command1_Click()
    
    GeneraBackup
    
End Sub

Private Sub Command2_Click()
    
    Call CopiaBkAlDropbox
    
End Sub

Private Sub Command3_Click()

    CopiaDeDropboxACarpetaBk

End Sub

Private Sub Command4_Click()

    RestauraDatos
    
End Sub

Private Sub Command6_Click()
    '
    n_file = FreeFile
    Open CarpetaDefault & "\confisincro.dat" For Random As #n_file Len = 128
    Put #n_file, 1, Text1.Text
    Put #n_file, 2, Text2.Text
    Put #n_file, 3, Text3.Text
    Put #n_file, 4, Text4.Text
        VALOR$ = "SUCU": If Option1.Value Then VALOR$ = "BSAS"
    Put #n_file, 5, VALOR$
    Put #n_file, 6, Text5.Text
    Close #n_file
    '
End Sub

Private Sub Command7_Click()
    
    Close: End
    
End Sub

Private Sub Form_Load()
    '
    Reintenta = False
    UltimoBackup = ""
    '
    On Error Resume Next
    n_file = FreeFile
    CarpetaDefault = "C:\FLEXOFT"
    Open CarpetaDefault & "\confisincro.dat" For Random As #n_file Len = 128
    FILA% = 0
    Do While Not EOF(n_file)
        FILA% = FILA% + 1
        Get #n_file, FILA%, mivariable$
        Select Case FILA%
            Case 1: Text1 = Trim$(mivariable): BASEDA = Text1
            Case 2: Text2 = Trim$(mivariable): RUTADB = Text2
            Case 3: Text3 = Trim$(mivariable): HORARIO = Text3
            Case 4: Text4 = Trim$(mivariable): RUTABK = Text4
            Case 5: ORIGEN = Trim$(mivariable)
            Case 6: Text5 = Trim$(mivariable): RUTASIS = Text5
        End Select
    Loop
    Close #n_file
    '
    If ORIGEN = "SUCU" Then
        Option2.Value = True
    End If
    '
    SUCURSAL = "NEUQUEN"
    '
    List1.Clear
    On Error Resume Next
    n_file = FreeFile
    Open CarpetaDefault & "\seguisincro.txt" For Input As #n_file
    If Err = 0 Then
        Do While Not EOF(n_file)
            Line Input #n_file, Linea
            List1.AddItem Linea
        Loop
        Close #n_file
    End If
    On Error GoTo 0
    
    Call ABRECONEXION
    
    'Timer1.Enabled = True
    
End Sub

Function GeneraBackup() As Integer

    GeneraBackup = 1
    Screen.MousePointer = 11
    On Error GoTo controlerrores
    If RUTABK$ <> "" And BASEDA$ <> "" Then
        'tmp_RUTABK$ = RUTABK$ & "\" & BASEDA$ & ".BAK"
        'sqlx$ = "BACKUP DATABASE " & BASEDA$ & " TO DISK='" & tmp_RUTABK$ & "' WITH INIT"
        'FLEXCONN.Execute sqlx$
    End If
    On Error GoTo 0
    Screen.MousePointer = 1
    Exit Function
    
controlerrores:
    Screen.MousePointer = 1
    GeneraBackup = 0 ' FALLO
    On Error GoTo 0
    Exit Function

End Function

Function CopiaBkAlDropbox() As Integer

    Screen.MousePointer = 11
    CopiaBkAlDropbox = 1
    On Error GoTo controlerrores
    If RUTABK$ <> "" And RUTADB$ <> "" And BASEDA$ <> "" Then
        tmp_RUTABK$ = RUTABK$ & "\" & BASEDA$ & ".BAK"
        tmp_RUTADB$ = RUTADB$ & "\" & BASEDA$ & ".BAK"
        Call FileCopy(tmp_RUTABK$, tmp_RUTADB$)
    End If
    On Error GoTo 0
    Screen.MousePointer = 1
    Exit Function
    
controlerrores:
    CopiaBkAlDropbox = 0
    Screen.MousePointer = 1
    On Error GoTo 0
    Exit Function

End Function

Function CopiaDeDropboxACarpetaBk() As Integer
    
    CopiaDeDropboxACarpetaBk = 1
    Screen.MousePointer = 11
    On Error GoTo controlerrores
    If RUTABK$ <> "" And RUTADB$ <> "" And BASEDA$ <> "" Then
        tmp_RUTABK$ = RUTABK$ & "\" & BASEDA$ & ".BAK"
        tmp_RUTADB$ = RUTADB$ & "\" & BASEDA$ & ".BAK"
        Call FileCopy(tmp_RUTADB$, tmp_RUTABK$)
    End If
    On Error GoTo 0
    Exit Function

controlerrores:
    CopiaDeDropboxACarpetaBk = 1
    Screen.MousePointer = 1
    On Error GoTo 0
    Exit Function
    
End Function

Function RestauraDatos()

    Screen.MousePointer = 11
    RestauraDatos = 1
    On Error GoTo controlerrores
    If RUTABK$ <> "" And BASEDA$ <> "" Then
'        tmp_RUTABK$ = RUTABK$ & "\" & BASEDA$ & ".BAK"
'        sqlx$ = "RESTORE DATABASE " & BASEDA$ & " FROM DISK ='" & tmp_RUTABK$ & "' WITH REPLACE"
'        FLEXCONN.Execute sqlx$
    End If
    On Error GoTo 0
    Screen.MousePointer = 1
    Exit Function
    
controlerrores:
    RestauraDatos = 0
    Screen.MousePointer = 1
    MsgBox Err.Description
    On Error GoTo 0
    Exit Function

End Function

Private Sub Timer1_Timer()

    AA$ = Format(Text3.Text, "hh:mm")
    BB$ = Format(Now, "hh:mm")
    BB1$ = Format(Now, "dd/mm/yy hh:mm")
    
    If AA$ = BB$ Or Reintenta = True Then
        Timer1.Enabled = False
        If UltimoBackup = BB1$ Then
            Timer1.Enabled = True
            Exit Sub
        End If
        If ORIGEN = "SUCU" Then
            If GeneraBackup = 1 Then
                'Call MsgBox("Backup OK")
                If CopiaBkAlDropbox = 1 Then
                    EscribeEnSeguimiento ("Backup generado el " & Str$(Now))
                Else
                    Reintenta = True
                    Call EscribeEnSeguimiento("Error al copiar al dropbox " & Str$(Now))
                End If
                Reintenta = False
                UltimoBackup = BB1$
            Else
                Reintenta = True
                Call EscribeEnSeguimiento("Error al generar backup " & Str$(Now))
            End If
        ElseIf ORIGEN = "BSAS" Then
            If CopiaDeDropboxACarpetaBk = 1 Then
                'Call MsgBox("Leer Backup OK")
                If RestauraDatos = 1 Then
                    'Call MsgBox("Restaurar OK")
                    Call EscribeEnSeguimiento("Backup restaurado el " & Str$(Now))
                    Reintenta = False
                    UltimoBackup = BB1$
                Else
                    Reintenta = True
                    Call EscribeEnSeguimiento("Error al restaurar " & Str$(Now))
                End If
            Else
                Reintenta = True
                Call EscribeEnSeguimiento("Error al copiar a backup  " & Str$(Now))
            End If
        End If
        Timer1.Enabled = True
    End If
    
End Sub

Sub EscribeEnSeguimiento(TEXTO As String)

    n_file = FreeFile
    Open CarpetaDefault & "\seguisincro.txt" For Append As #n_file
    TXT$ = TEXTO
    Print #n_file, TXT$
    Close #n_file
    List1.AddItem TXT$

End Sub

Sub ABRECONEXION()
   '
   Exit Sub
   '
   On Error Resume Next
   ESTADODECONEXION% = FLEXCONN.State
   If ESTADODECONEXION% = 1 Then
      Exit Sub
   End If
   '
   CONNSQL$ = Trim$(LOADFILE$(CarpetaDefault + "\CONNSQLS.DRV"))
'MsgBox CONNSQL$
   If CONNSQL$ <> "" Then
      ' VERIFICA LA CONEXION
      Set FLEXCONN = New ADODB.Connection
      FLEXCONN.ConnectionString = CONNSQL$   ' "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=True;Data Source=FLEXSERVER;Initial Catalog=" + conecta$ + "; Connect Timeout=999"
      On Error Resume Next
      FLEXCONN.Open
      If Err Then
         MsgBox "Imposible Abrir Conexión a Servidor SQL con " + CONNSQL$, vbCritical
         Close: End
      End If
      On Error GoTo 0
      '
      FLEXCONN.Execute ("SET DATEFORMAT dmy")
      FLEXCONN.CommandTimeout = 240 '120
      TIBASE$ = "SQL"
      TXT$ = CONNSQL$
      PINI% = InStr(UCase$(TXT$), "DATA SOURCE=") + 12
      TXT1$ = Mid$(TXT$, PINI%)
      PFIN% = InStr(TXT1$, ";")
      SERNA$ = Left$(TXT1$, PFIN% - 1)
      PINI% = InStr(UCase$(TXT$), "INITIAL CATALOG=") + 16
      TXT1$ = Mid$(TXT$, PINI%)
      PFIN% = InStr(TXT1$, ";")
      BINIC$ = Left$(TXT1$, PFIN% - 1)
      '
      BASEACTIVA$ = BINIC$
      TIBASE$ = "SQL"
      GoTo 50
   Else
      MsgBox "Archivo de Ruta de Conexión a Servidor Sql " & vbCrLf & "(CONNSQL.DRV) Inexistente o Vacio ", vbCritical
      Close: End
   End If
   '
   On Error Resume Next
   XX$ = CONN_SQLS.Label6.Caption
   If Err Then
      On Error GoTo 0
      MsgBox "No se han Establecido los Parámetros\de Conexión al Servidor SQL-Server."
   End If
   On Error GoTo 0
   CONN_SQLS.Show 1
   MsgBox "Re-Inicie Ahora la Aplicación."
   '
50
'   '
End Sub

Function LOADFILE$(FINA$)
   '
   LFX$ = ""
   SCMA = Screen.MousePointer
   Screen.MousePointer = 11
   '
   On Error GoTo 97
   LOPUCO& = FileLen(FINA$)
   NX% = FILEOPEN%(FINA$, 41, 2048)
   LOPUCF& = LOF(NX%)
   If LOPUCF& > LOCUPO& Then LOPUCO& = LOPUCF&
   '
   LFX$ = String$(LOPUCO&, 0)
   Get #NX%, 1, LFX$
   GoTo 99
   '
97 Resume 98
98 LFX$ = ""
   '
99 If NX% > 0 Then Close #NX%
   On Error GoTo 0
   Screen.MousePointer = SCMA
   LOADFILE$ = LFX$
   '
End Function


Function FILEOPEN%(ARCHX$, MODO%, LOREGG%)
    '
    MODACCE$ = "ASIGNACION NUMERO DE ARCHIVO"
    On Error GoTo 99
    FILI% = FreeFile
    Select Case MODO%
      Case 0
        MODACCE$ = "ACCESO ALEATORIO"
        Open ARCHX$ For Random Access Read Write Shared As #FILI% Len = LOREGG%
        GoTo 89
      Case 10
        MODACCE$ = "ACCESO ALEATORIO SOLO LECTURA"
        Open ARCHX$ For Random Access Read Shared As #FILI% Len = LOREGG%
        GoTo 89
      Case 1
        MODACCE$ = "LECTURA SECUENCIAL"
        Open ARCHX$ For Input Shared As #FILI%
        GoTo 89
      Case 2
        MODACCE$ = "ESCRITURA SECUENCIAL"
        Open ARCHX$ For Output As #FILI%
        GoTo 89
      Case 3
        MODACCE$ = "SALIDA SECUENCIAL INCREMENTAL"
        Open ARCHX$ For Append As #FILI%
        GoTo 89
      Case 4
        MODACCE$ = "ACCESO BINARIO"
        Open ARCHX$ For Binary Access Read Write Shared As #FILI%
        GoTo 89
      Case 41
        MODACCE$ = "ACCESO BINARIO SOLO LECTURA" ' LOADFILE
        Open ARCHX$ For Binary Access Read Shared As #FILI%
        GoTo 89
     Case 9
        MODACCE$ = "ACCESO ALEATORIO SOLO LECTURA"
        Open ARCHX$ For Random Access Read Shared As #FILI% Len = LOREGG%
        GoTo 89
    End Select
    FILI% = 0
    '
89  If FILI% > 0 Then
        On Error GoTo 0
        FILEOPEN% = FILI%
        Exit Function
    End If
    GoTo 100
    '
99  ERNU$ = Trim$(Str$(Err.Number))
    '
100 BRXX$ = REPLACAR$(ARCHX$, "\", "/")
    If Mid$(BRXX$, 2, 1) <> ":" Then
      If Left$(BRXX$, 2) <> "//" Then
        If Mid$(BRXX$, 2, 1) <> ":" Then
          BRXX$ = Left$(CurDir, 2) + BRXX$
        End If
      End If
    End If
    MENSA$ = "Se ha Producido un Error (" + ERNU$ + ") en la Apertura para\"
    MENSA$ = MENSA$ + MODACCE$ + " del Siguiente Archivo:\  \"
    MENSA$ = MENSA$ + BRXX$ + "\  \"
    '
    MENSA$ = MENSA$ + "Investigue las Siguientes Causas Posibles:\  \"
    MENSA$ = MENSA$ + "   1.- Problemas de Red - No hay Acceso a U.L. '" + Left$(BRXX$, 2) + "'.\"
    MENSA$ = MENSA$ + "   2.- Ruta de Acceso Inexistente.\"
    MENSA$ = MENSA$ + "   3.- Archivo de Solo Lectura u Oculto.\"
    MENSA$ = MENSA$ + "   4.- Archivo Inexistente (LECTURA SECUENCIAL).\"
    MENSA$ = MENSA$ + "   5.- Archivo no Compartido en Uso por Otra Aplicación.\"
    MENSA$ = MENSA$ + "   \Consulte a su Soporte de Sistemas.\"
    '
    c$ = (MENSA$)
    RESPU% = MsgBox(c$, 53, "Error de Apertura de Archivo")
    If RESPU% = 4 Then
        Resume
    End If
    Close: End
    '
End Function

Function REPLACAR$(A$, b$, c$)
    A1$ = A$
    '
    If Len(b$) > 0 Then
10        XL& = InStr(A1$, b$)   ' Reemplazado XL% por XL& para evitar desbordamiento - EPB 16/01/11
          If XL& > 0 Then
              A1$ = Left$(A1$, XL& - 1) + c$ + Mid$(A1$, XL& + Len(b$))
              GoTo 10
          End If
    End If
    '
    REPLACAR$ = A1$
    A1$ = ""
    '
End Function

Private Sub Timer2_Timer()

    If ORIGEN = "SUCU" Then
        On Error Resume Next
        ' RECIBE ATRASUCU
        ATRASUCU_LO$ = LOADFILE(RUTASIS & "\ATRASUC.DAT")
        ATRASUCU_DB$ = LOADFILE(RUTADB & "\ATRASUC.DAT")
        If ATRASUCU_DB$ <> ATRASUCU_LO$ And ATRASUCU_DB$ <> "" Then
            Call FileCopy(RUTADB & "\ATRASUC.DAT", RUTASIS & "\ATRASUC.DAT")
            If Err.Number = 0 Then
                EscribeEnSeguimiento ("Recepcion de ATRASUCU " & Str$(Now))
            Else
                EscribeEnSeguimiento ("Fallo al Leer ATRASUCU " & Str$(Now))
            End If
        End If
        '
        ' ENVIA ATRACEN
        ATRACEN_LO$ = LOADFILE(RUTASIS & "\ATRACEN.DAT")
        ATRACEN_DB$ = LOADFILE(RUTADB & "\ATRACEN.DAT")
        If ATRACEN_DB$ <> ATRACEN_LO$ And ATRACEN_LO$ <> "" Then
            Call FileCopy(RUTASIS & "\ATRACEN.DAT", RUTADB & "\ATRACEN.DAT")
            If Err.Number = 0 Then
                EscribeEnSeguimiento ("Envio de ATRACEN " & Str$(Now))
            Else
                EscribeEnSeguimiento ("Fallo al Leer ATRACEN " & Str$(Now))
            End If
        End If
        '
        ' **********************************************************************************
        '
        ' RECIBE ARCHIVOS DEL MANEJO DE LAS SOLICITUDES DESDE LA SUCURSALES
        sName = Dir(RUTADB & "\ARCHIVO*.TXT", vbNormal)
        While Len(sName) > 0
            ArchivoOrigen = RUTADB & "\" & sName
            ArchivoDestino = RUTASIS & "\" & sName
            SetAttr ArchivoOrigen, vbNormal
            Call FileCopy(ArchivoOrigen, ArchivoDestino)
            If Err.Number = 0 Then
                Kill ArchivoOrigen
                EscribeEnSeguimiento ("Recibe de " & sName & " - " & Str$(Now))
            End If
            sName = Dir
        Wend
        
        ' ENVIA ARCHIVOS DEL MANEJO DE LAS SOLICITUDES DESDE LA SUCURSALES
        sName = Dir(RUTASIS & "\OCOMPRA*.TXT", vbNormal)
        While Len(sName) > 0
            ArchivoOrigen = RUTASIS & "\" & sName
            ArchivoDestino = RUTADB & "\" & sName
            SetAttr ArchivoOrigen, vbNormal
            Call FileCopy(ArchivoOrigen, ArchivoDestino)
            If Err.Number = 0 Then
                Kill ArchivoOrigen
                EscribeEnSeguimiento ("Envio de " & sName & " - " & Str$(Now))
            End If
            sName = Dir
        Wend
        '
        On Error GoTo 0
        '
    ElseIf ORIGEN = "BSAS" Then
        On Error Resume Next
        If SUCURSAL = "NEUQUEN" Then
            SUCURSAL = "COMODORO"
        Else
            SUCURSAL = "NEUQUEN"
        End If
        
        ' ENVIA ATRASUCU
        ATRASUCU_LO$ = LOADFILE(RUTASIS & "\" & SUCURSAL & "\ATRASUC.DAT")
        ATRASUCU_DB$ = LOADFILE(RUTADB & "\" & SUCURSAL & "\ATRASUC.DAT")
        If ATRASUCU_DB$ <> ATRASUCU_LO$ And ATRASUCU_LO$ <> "" Then
            Call FileCopy(RUTASIS & "\" & SUCURSAL & "\ATRASUC.DAT", RUTADB & "\" & SUCURSAL & "\ATRASUC.DAT")
            If Err.Number = 0 Then
                EscribeEnSeguimiento ("Envio de ATRASUCU (" & SUCURSAL & ") " & Str$(Now))
            Else
                EscribeEnSeguimiento ("Fallo al Enviar ATRASUCU (" & SUCURSAL & ") " & Str$(Now))
            End If
        End If
        '
        ' RECIBE ATRACEN
        ATRACEN_LO$ = LOADFILE(RUTASIS & "\" & SUCURSAL & "\ATRACEN.DAT")
        ATRACEN_DB$ = LOADFILE(RUTADB & "\" & SUCURSAL & "\ATRACEN.DAT")
        If ATRACEN_DB$ <> ATRACEN_LO$ And ATRACEN_DB$ <> "" Then
            Call FileCopy(RUTADB & "\" & SUCURSAL & "\ATRACEN.DAT", RUTASIS & "\" & SUCURSAL & "\ATRACEN.DAT")
            If Err.Number = 0 Then
                EscribeEnSeguimiento ("Recepcion de ATRACEN (" & SUCURSAL & ") " & Str$(Now))
            Else
                EscribeEnSeguimiento ("Fallo al Leer ATRACEN (" & SUCURSAL & ") " & Str$(Now))
            End If
        End If
        '
        ' **********************************************************************************
        '
        ' RECIBE ARCHIVOS DEL MANEJO DE LAS SOLICITUDES DESDE LA SUCURSALES
        sName = Dir(RUTADB & "\" & SUCURSAL & "\OCOMPRA*.TXT", vbNormal)
        While Len(sName) > 0
            ArchivoOrigen = RUTADB & "\" & SUCURSAL & "\" & sName
            ArchivoDestino = RUTASIS & "\" & SUCURSAL & "\" & sName
            SetAttr ArchivoOrigen, vbNormal
            Call FileCopy(ArchivoOrigen, ArchivoDestino)
            If Err.Number = 0 Then
                Kill ArchivoOrigen
                EscribeEnSeguimiento ("Envio de " & sName & " (" & SUCURSAL & ") " & Str$(Now))
            End If
            sName = Dir
        Wend
        
        ' ENVIA ARCHIVOS DEL MANEJO DE LAS SOLICITUDES DESDE LA SUCURSALES
        sName = Dir(RUTASIS & "\" & SUCURSAL & "\ARCHIVO*.TXT", vbNormal)
        While Len(sName) > 0
            ArchivoOrigen = RUTASIS & "\" & SUCURSAL & "\" & sName
            ArchivoDestino = RUTADB & "\" & SUCURSAL & "\" & sName
            SetAttr ArchivoOrigen, vbNormal
            Call FileCopy(ArchivoOrigen, ArchivoDestino)
            If Err.Number = 0 Then
                Kill ArchivoOrigen
                EscribeEnSeguimiento ("Envio de " & sName & " (" & SUCURSAL & ") " & Str$(Now))
            End If
            sName = Dir
        Wend
        '
        On Error GoTo 0
        '
    End If
    
End Sub
