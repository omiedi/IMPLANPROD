VERSION 5.00
Begin VB.Form FSETUP00 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración de Funcionamiento"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6255
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   6255
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command222 
      Caption         =   "Depurar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5250
      Picture         =   "FSETUP00.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Depurar Movimientos Viejos de B.Datos Movimientos de Stock"
      Top             =   4095
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Clave"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5250
      Picture         =   "FSETUP00.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3150
      Width           =   855
   End
   Begin VB.CommandButton Command29 
      Caption         =   "Setup"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5250
      Picture         =   "FSETUP00.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Configuración General de Funcionamiento"
      Top             =   2205
      Width           =   855
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Control de Proceso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   210
      TabIndex        =   11
      Top             =   3570
      Width           =   4845
      Begin VB.CommandButton Command26 
         Caption         =   "Sistema"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1890
         TabIndex        =   20
         Top             =   1575
         Width           =   1170
      End
      Begin VB.CommandButton Command25 
         Caption         =   "Ventanas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   315
         TabIndex        =   19
         Top             =   1575
         Width           =   1170
      End
      Begin VB.CommandButton Command24 
         Caption         =   "Ordenador"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1890
         TabIndex        =   18
         Top             =   1050
         Width           =   1170
      End
      Begin VB.CommandButton Command23 
         Caption         =   "Objetos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3360
         TabIndex        =   17
         Top             =   1575
         Width           =   1170
      End
      Begin VB.CommandButton Command22 
         Caption         =   "Formularios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1890
         TabIndex        =   16
         Top             =   525
         Width           =   1170
      End
      Begin VB.CommandButton Command21 
         Caption         =   "Filtradores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   315
         TabIndex        =   15
         Top             =   1050
         Width           =   1170
      End
      Begin VB.CommandButton Command20 
         Caption         =   "Estructuras"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   315
         TabIndex        =   14
         Top             =   525
         Width           =   1170
      End
      Begin VB.CommandButton Command19 
         Caption         =   "Listadores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3360
         TabIndex        =   13
         Top             =   525
         Width           =   1170
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Ficheros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3360
         TabIndex        =   12
         Top             =   1050
         Width           =   1170
      End
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Terminar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5250
      Picture         =   "FSETUP00.frx":0A56
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   315
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ayuda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5250
      Picture         =   "FSETUP00.frx":0BA0
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1260
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Avanzadas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1590
      Left            =   210
      TabIndex        =   1
      Top             =   1890
      Width           =   4845
      Begin VB.CommandButton Command9 
         Caption         =   "Formularios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   315
         TabIndex        =   5
         Top             =   1050
         Width           =   2010
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Diseño de Menúes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2520
         TabIndex        =   4
         Top             =   1050
         Width           =   2010
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Multiempresa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2520
         TabIndex        =   3
         Top             =   525
         Width           =   2010
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Carpetas de Trabajo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   315
         TabIndex        =   2
         Top             =   525
         Width           =   2010
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Administración"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1590
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   4845
      Begin VB.CommandButton Command27 
         Caption         =   "Empresa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   945
         TabIndex        =   21
         Top             =   525
         Width           =   1275
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Usuarios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2730
         TabIndex        =   8
         Top             =   525
         Width           =   1275
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Backup"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   945
         TabIndex        =   7
         Top             =   1050
         Width           =   1275
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Auditoría"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2730
         TabIndex        =   6
         Top             =   1050
         Width           =   1275
      End
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   4695
      Picture         =   "FSETUP00.frx":0EAA
      Top             =   5145
      Width           =   2010
   End
End
Attribute VB_Name = "FSETUP00"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command10_Click()
    '
    QWR% = USNBR% Mod 100
    If QWR% <> 1 Then
        Call COMUNI("La Opcion Elegida\Requiere Privilegios de Supervisor")
        Exit Sub
    End If
    '
    OPX% = COMALTER%("Elija la Función de Auditoría:\\Utilizacion\Mensajes Emergentes\Bloqueos Activos\Histórico de Bloqueos")
    If OPX% = 1 Then
         GoSub AUDITO
       ElseIf OPX% = 2 Then
         Call SELECHO("SELFECHA")
         FFXX$ = VALACT1$("SELFECHA")
         If FFXX$ = "" Then Exit Sub
         FFII% = FECHANUM(FFXX$)
         JJ& = 0
         Screen.MousePointer = 11
         For U& = 1 To ULTREG&("MENSAJES")
            XX$ = REGLEIDO$("MENSAJES", U&)
            If CVI(Mid$(XX$, 5, 2)) = FFII% Then
               JJ& = JJ& + 1
               Call GRAREG("MENSAFIL", XX$, JJ&)
            End If
         Next U&
         Call SETULTREG("MENSAFIL", JJ&)
         Call CIERRARCH("MENSAFIL")
         Call CIERRARCH("MENSAJES")
         Screen.MousePointer = 1
         Call CONECRUN("*LIMENSA", "Listado de Mensajes Emergentes")
       ElseIf OPX% = 3 Then
         Call MONIBLOQ
         Exit Sub
         '****************************************
         LU$ = LUDAT$
10       Screen.MousePointer = 11
         Call BLANARCH("!BLOQAR")
         For U& = 1 To ULTREG&("BLOQAR")
            XXX$ = REGLEIDO$("BLOQAR", U&)
            If TRIM$(Mid$(XXX$, 13, 116)) <> "" Then
               Call REGAPP("!BLOQAR", XXX$)
            End If
         Next U&
         Call CIERRARCH("BLOQAR")
         Call CIERRARCH("!BLOQAR")
         Screen.MousePointer = 1
         VTB% = VENTABU%("BLOQAR")
         If VTB% >= 0 Then GoTo 10
       ElseIf OPX% = 4 Then
         GoSub SECUBLOQ
    End If
    '
    Exit Sub
    '
    '
AUDITO:
   If EXISTE%("C:\FLEXOFT\WORDPAD.EXE") = 1 Then
        Shell "C:\FLEXOFT\WORDPAD.EXE " + LUDAT$ + "MONIEXEC.DAT", vbMaximizedFocus
      Else
        MENSA$ = "Imposible Mostrar Auditoría de Uso."
        MENSA$ = MENSA$ + "\  \Falta Utilitario 'WORDPAD.EXE'"
        MENSA$ = MENSA$ + "\en Directorio 'C:/FLEXOFT'."
        MENSA$ = MENSA$ + "\   \Instale este Programa Utilitario\y Repita la Consulta."
        Call COMUNI(MENSA$)
   End If
Return
    '
    '
SECUBLOQ:
   HOO% = HOY(HO$)
   VDEF$ = MKI$(HOO%) + MKI$(HOO%)
   OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
   If OBX$ = "" Then Exit Sub
   DES% = CVI(Left$(OBX$, 2))
   HAS% = CVI(Mid$(OBX$, 3, 2))
   If DES% > HAS% Then Exit Sub
   '
   Screen.MousePointer = 11
   On Error Resume Next
   Kill LUDAT$ + "SEBLOFIL.DAT"
   On Error GoTo 0
   Dim TSX1 As String * 128
   JJ& = 0
   NX1% = FreeFile
   Open LUDAT$ + "SECUBLOQ.DAT" For Random Access Read Write As #NX1% Len = 128
   NX2% = FreeFile
   Open LUDAT$ + "SEBLOFIL.DAT" For Random Access Read Write As #NX2% Len = 128
   For U& = 1 To LOF(NX1%) / 128
     Get #NX1%, U&, TSX1$
     FXYZ% = FECHANUM(Mid$(TSX1$, 25, 8))
     If FXYZ% >= DES% Then
       If FXYZ% <= HAS% Then
         JJ& = JJ& + 1
         Put #NX2%, JJ&, TSX1$
       End If
     End If
   Next U&
   Close #NX1%: Close #NX2%
   Screen.MousePointer = 1
   '
   If EXISTE%("C:\FLEXOFT\WORDPAD.EXE") = 1 Then
        Shell "C:\FLEXOFT\WORDPAD.EXE " + LUDAT$ + "SEBLOFIL.DAT", vbMaximizedFocus
      Else
        MENSA$ = "Imposible Mostrar Histórico de Bloqueos."
        MENSA$ = MENSA$ + "\  \Falta Utilitario 'WORDPAD.EXE'"
        MENSA$ = MENSA$ + "\en Directorio 'C:/FLEXOFT'."
        MENSA$ = MENSA$ + "\   \Instale este Programa Utilitario\y Repita la Consulta."
        Call COMUNI(MENSA$)
   End If
Return
'
End Sub

Private Sub Command17_Click()
    Call CIERRARCH("*.*")
    Call COMUNI("Re-Configuración Completa.\Re-Inicie Ahora la Aplicación.")
    Call FINAL("")
End Sub

Private Sub Command11_Click()
    Call CONECRUN("#FLEXRCV", "Copias de Seguridad")
End Sub


Private Sub Command16_Click()
    '
    If DERACCE%("ACTUSER", "Usuarios del Sistema FLEXOFT") = 2 Then GoTo 20
    '
    QWR% = USNBR% Mod 100
    If QWR% <> 1 Then
        Call COMUNI("La Opcion Elegida\Requiere Privilegios de Supervisor")
        Exit Sub
    End If
    '
20  MODO% = ALTERNA%("A-B-M Usuarios del Sistema\Derechos de Acceso por Usuario")
    If MODO% = 1 Then
          Call ACTUSER
        ElseIf MODO% = 2 Then
          Call ACDERACC
    End If
    '
End Sub

Private Sub Command1_Click()
    Call COMUNI("Ayuda en Línea no Disponible")
End Sub

Private Sub Command18_Click()
    Call CONNDOS
    Call CONECRUN("#PARAPLA", "Definición de Ficheros Planos")
End Sub

Private Sub Command19_Click()
    Call CONNDOS
    Call CONECRUN("#PARALIS", "Definicion de Listadores Tabulados")
End Sub

Private Sub Command2_Click()
   Screen.MousePointer = 11
   FIN& = ULTREG&("PDOCLST")
   For U& = 73740 To FIN&
     XX$ = REGLEIDO$("PDOCLST", U&)
     Call REPLA(XX$, MKS$(U&), 1, 4)
     If U& >= 73740 Then
       If U& <= 73742 Then
         Call REPLA(XX$, "Documento Numero " + TRIM$(Str$(U&)), 5, 124)
       End If
     End If
     Call GRAREG("PDOCLST", XX$, U&)
   Next U&
   Call CIERRARCH("PDOCLST")
   Call FRESHECHO("PDOCLST")
   '
   FIN& = ULTREG&("PDOCSPL")
   TTXX$ = ""
   For U& = 1470000 To FIN&
     XX$ = REGLEIDO$("PDOCSPL", U&)
     IU& = CVS(Left$(XX$, 4))
     If IU& < 10000 Then
       IU& = IU& + 73742
       Call REPLA(XX$, MKS$(IU&), 1, 4)
       Call GRAREG("PDOCSPL", XX$, U&)
     End If
   Next U&
   Call CIERRARCH("PDOCSPL")
   Screen.MousePointer = 1
   Call COMUNI("Recuperacion Completada")
End Sub

Private Sub Command20_Click()
    Call CONNDOS
    Call CONECRUN("#PARASTR", "Definición Estructura de Archivos")
End Sub

Private Sub Command21_Click()
    Call CONNDOS
    Call CONECRUN("#PARAFIL", "Procedimientos de Filtrado")
End Sub

Private Sub Command22_Click()
    Call CONNDOS
    Call CONECRUN("#PARAMED", "Definición de Formularios Tabulados")
End Sub

 Sub Command222_Click()
  Call OPTRAJUST.Command17_Click
End Sub

Private Sub Command23_Click()
    If EXISTE%("OBJEDIT.PRF") < 1 Then
        Call EXTRACT("OBJEDIT.PRF")
    End If
    Call CONNDOS
    Call CONECRUN("#PARAOBJ", "Definicion de Objetos Planos")
End Sub

Private Sub Command24_Click()
    Call CONNDOS
    Call CONECRUN("#PARASOR", "Definicion de Modulo de Ordenamiento")
End Sub

Private Sub Command25_Click()
    Call CONNDOS
    Call CONECRUN("#PARAVEN", "")
End Sub

Private Sub Command26_Click()
    SVLC$ = SERIVOL$("C:")
    If SVLC$ <> "081D-0BD9" Then      ' MAQUINA ERNESTO
       If SVLC$ <> "2560-10F0" Then   ' NOTEBOOK
          Call CARDASYS
          Exit Sub
       End If
    End If
    '
    MODO% = ALTERNA%("Diskette de Habilitación\Identificador del Sistema")
    If MODO% < 1 Or MODO% > 2 Then
        Exit Sub
    End If
    If MODO% = 2 Then
        Call CARDASYS
        Exit Sub
    End If
    '
    Call COMUNI("Coloque un Diskette Formateado y en Blanco")
    SVLA$ = TRIM$(SERIVOL$("A:"))
    If SVLA$ = "" Then
          GoTo 99
        ElseIf SVLA$ = "0000-0000" Then
          Call MENSERR(24, "Diskette no Correctamente Formateado")
          Exit Sub
    End If
    '
    OBX$ = OBJPLA$("HABILITASYS", "")
    If OBX$ = "" Then
        Exit Sub
    End If
    '
    SVL$ = Left$(OBX$, 9)
    FLIM% = CVI(Mid$(OBX$, 14, 2))
    '
    On Error GoTo 98
    NY% = FreeFile
    Open "A:FLEXCHK.DRV" For Random Access Read Write Shared As #NY% Len = 127
    Dim CHKY As String * 127
    Randomize Timer
    For KK% = 1 To 8196
        CHKZ$ = String$(127, 0)
        For KL% = 1 To 127
           PPX% = Int(Rnd * 255)
           Mid$(CHKZ$, KL%, 1) = Chr$(PPX%)
        Next KL%
        LSet CHKY$ = CHKZ$
        Put #NY%, KK%, CHKY$
    Next KK%
    '
    Get #NY%, 1953, CHKY$
    CHKZ$ = CHKY$
    SVL1$ = ENCRIPTA$(SVL$)
    Call REPLA(CHKZ$, SVL1$, 2, 9)
    '
    SVLA$ = TRIM$(SERIVOL$("A:"))
    SVLA1$ = ENCRIPTA$(SVLA$)
    Call REPLA(CHKZ$, SVLA1$, 73, 9)
    '
    Call REPLA(CHKZ$, MKI$(FLIM%), 92, 2)
    '
    PPXA% = 0
    For KK% = 1 To 94
        PPXA% = PPXA% + Asc(Mid$(CHKZ$, KK%, 1)) - 128 '
    Next KK%
    Call REPLA(CHKZ$, MKI$(PPXA%), 95, 2)
    '
    LSet CHKY$ = CHKZ$
    Put #NY%, 1953, CHKY$
    Close #NY%
    On Error GoTo 0
    Call COMUNI("Diskette Completo")
    Exit Sub
    '
98  Resume 99
99  On Error GoTo 0
    Call MENSERR(24, "Error de Generacion de Diskette")
    Exit Sub
    '
End Sub

Private Sub Command27_Click()
    DATEMPRE.Show 1
    Call FRESHALL
End Sub

Private Sub Command29_Click()
    FLEXSETUP.Show 1
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    '
    Dim F0 As String * 64
    NX5% = FILEOPEN%("USER.DAT", 0, 64)
    If LOF(NX5%) < 1 Then
       Call MENSERR(24, "Falta Definir Usuarios FLEXOFT")
       GoTo 99
    End If
    '
    JJ% = USNBR% Mod 100
    If JJ% > 0 Then
       If JJ% <= LOF(NX5%) / 64 Then
          On Error GoTo 90
          Get #NX5%, JJ%, F0$
          IDENT$ = Mid$(F0$, 31, 12): If Left$(IDENT$, 1) = "@" Then IDENT$ = DESENCRI$(Mid$(F0$, 32, 11)) + " "
          DENO$ = Left$(F0$, 30): If Left$(DENO$, 1) = "@" Then DENO$ = DESENCRI$(Mid$(F0$, 2, 29)) + " "
          CLAV$ = Mid$(F0$, 43, 12): If Left$(CLAV$, 1) = "@" Then CLAV$ = DESENCRI$(Mid$(F0$, 44, 11)) + " "
          USU$ = TRIM$(IDENT$) + " - " + TRIM$(DENO$)
          VDEF$ = Space$(72)
          Call REPLA(VDEF$, USU$, 1, 48)
10        Call REPLA(VDEF$, String$(Len(CLAV$), "*"), 49, 12)
          OBX$ = OBJPLA$("CAMBICLAVE", VDEF$)
          If OBX$ = "" Then GoTo 99
          VD11$ = TRIM$(Mid$(OBX$, 61, 12))
          If Len(VD11$) < 1 Then
             Call MENSERR(24, "Clave no Válida")
             GoTo 10
          End If
          For K% = 1 To Len(VD11$)
             K1% = Asc(Mid$(VD11$, K%, 1))
             If K1% < 48 Or K1% > 90 Then
                Call MENSERR(24, "Clave no Válida")
                GoTo 10
             End If
             If K1% > 57 And K1% < 65 Then
                Call MENSERR(24, "Clave no Válida")
                GoTo 10
             End If
          Next K%
          CLAVE$ = Left$(VD11$ + Space$(12), 12)
          '
          Get #NX5%, JJ%, F0$
          DE$ = ENCRIPTA$(CLAVE$)
          CLAVE$ = "@" + DE$
          Call REPLA(F0$, CLAVE$, 43, 12)
          Put #NX5%, JJ%, F0$
       End If
    End If
    GoTo 99
    '
90  Resume 91
91  Call MENSERR(24, "Imposible Registrar Nueva Clave.")
    '
99  On Error GoTo 0
    Close #NX5%
    Command5.Enabled = True
    '
End Sub

Private Sub Command6_Click()
    Call DIRETRA
End Sub

Private Sub Command7_Click()
   Command7.Enabled = False
   Call CARMULTIEM
   Command7.Enabled = True
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   Call CONECRUN("#FLEXMNU", "Configuracion de Menúes")
   Command8.Enabled = True
End Sub

Private Sub Command9_Click()
    '
    Command9.Enabled = False
    ACONEC$ = "FLFORMS"
10  Call CONECRUN(ACONEC$, "Formularios de Impresion")
    Command9.Enabled = True
    '
End Sub

Private Sub Form_Load()
    '
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    COMMA$ = TRIM$(Command$)
    PPX% = InStr(COMMA$, "#" + JOBID$ + "#")
    If PPX% > 0 Then
        COMMA$ = Left$(COMMA$, PPX% - 1) + Mid$(COMMA$, PPX% + 10)
    End If
    '
    If APLINVO$ <> "" Then
       On Error Resume Next
       If APLINVO$ = "ANACOSTO" Then
           OPANACOS.Show 1
         ElseIf APLINVO$ = "OPCLAVEN" Then
           OPCLAV00.Show 1
         ElseIf APLINVO$ = "OPFACTUR" Then
           OPFACTU00.Show 1
         ElseIf APLINVO$ = "OPRECIBO" Then
           OPRECIB00.Show 1
         ElseIf APLINVO$ = "OPOCOMPRA" Then
           OPOCOMPRA.Show 1    ' Power-Up Industrial
         ElseIf APLINVO$ = "SETUPCOM" Then
           SETUPCOM.Show 1     ' ORCOMP03
         ElseIf APLINVO$ = "OPTRAJUST" Then
           Screen.MousePointer = 11
           OPTRAJUST.Show 1
         ElseIf APLINVO$ = "OPFABR00" Then
           OPFABR00.Show 1
         ElseIf APLINVO$ = "OPAMACLI" Then
           OPAMACLI.Show 1
         ElseIf APLINVO$ = "OPMAST00" Then
           OPMAST00.Show 1
         ElseIf APLINVO$ = "OPPROMRP" Then
           OPPROMRP.Show 1
         ElseIf APLINVO$ = "OPIRMP00" Then
           OPIRMP00.Show 1
         ElseIf APLINVO$ = "SETUPAEN" Then
           OPAENTRE.Show 1
       End If
       On Error GoTo 0
       Call COMUNI("Re-Configuración Completa.\Re-Inicie Ahora la Aplicación.")
       Call FINAL("")
    End If
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    LCB$ = "C"
10  Call GRARAM(97, LCB$)
    On Error GoTo 800
    N1% = FreeFile
    Open LCB$ + ":\COMMAND.COM" For Input Shared As #N1%
    If N1% > 0 Then Close #N1%: N1% = 0
    On Error GoTo 0
    '
    If LEERAM$(97, 1) = LCB$ Then GoTo 11
    If LCB$ = "A" Then
        LCB$ = "C"
        Call GRARAM(97, LCB$)
        GoTo 11
    End If
    LCB$ = "A": GoTo 10
    '
11  LUBOOT$ = LCB$
    Exit Sub
    '
    '
800 LCB$ = LEERAM$(97, 1)
    If LCB$ = "C" Then
            Call GRARAM(97, "A")
          Else
            Call GRARAM(97, "C")
    End If
    Resume Next
    '
End Sub

Sub MONIBLOQ()              ' monitoreo de bloqueos
   '
   Dim AA As String * 128
   Dim BB As String * 128
   Static ARCHNAME$(2048), UACTU#(2048), CANARBLO%
   FECHOY& = Int(Now)
   '
10 Screen.MousePointer = 11
   LU$ = LUDAT$
   FORMALTER.File1.Path = LUDAT$
   FORMALTER.File1.Pattern = "*.LCK"
   '
12 Screen.MousePointer = 11
   JJ& = 0
   For IABLO% = 1 To FORMALTER.File1.ListCount
      B1$ = TRIM$(UCase$(FORMALTER.File1.List(IABLO% - 1)))
      If EXISTE%(LU$ + B1$) = 1 Then
         On Error GoTo 15
         NK% = FreeFile
         Open LU$ + B1$ For Random Access Read Write Lock Write As #NK% Len = 128
         Get #NK%, 1, AA$
         Put #NK%, 1, AA$
         Close #NK%: NK% = 0
         On Error GoTo 0
         GoTo 19
         '
15       Resume 16
16       On Error GoTo 0
         Close #NK%: NK% = 0
         NK% = FreeFile
         Open LU$ + B1$ For Random Access Read Shared As #NK% Len = 128
         Get #NK%, 1, AA$
         JJ& = JJ& + 1
         Call GRAREG("!BLOQAR", AA$, JJ&)
         Close #NK%: NK% = 0
      End If
19 Next IABLO%
   '
   Call SETULTREG("!BLOQAR", JJ&)
   Call CIERRARCH("!BLOQAR")
   Screen.MousePointer = 1
   VTB% = VENTABU%("BLOQAR")
   If VTB% >= 0 Then GoTo 12
   '
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Cancel = 1
End Sub

Sub CARMULTIEM()
   '
   TTXX$ = "En la Presente Versión\las Opciones Multiempresa Pueden Configurarse\Unicamente Bajo Supervisión Directa de FLEXOFT"
   If COMALTER%(TTXX$ + "\\Abandonar\Continuar") = 2 Then
     If CLACONTRA% = 1 Then
       Call BLANARCH("!CARMULTI")
       NX% = FILEOPEN("FLEXMUL.DRV", 0, 79)
       Dim FMX As String * 79
       '
       For I% = 1 To LOF(NX%) / 79
         Get #NX%, I%, FMX$
         If Left$(FMX$, 77) <> String$(77, 0) Then
           SUCAR% = 0
           For J% = 1 To 77
             SUCAR% = SUCAR% + Asc(Mid$(FMX$, J%, 1))
           Next J%
           If CVI(Mid$(FMX$, 78, 2)) = SUCAR% Then
             RASO$ = DESENCRI$(Left$(FMX$, 30))
             PACCE$ = DESENCRI$(Mid$(FMX$, 32, 46))
             Z$ = Space$(96)
             Call REPLA(Z$, RASO$, 1, 30)
             Call REPLA(Z$, PACCE$, 31, 46)
             Call REGAPP("!CARMULTI", Z$)
           End If
         End If
       Next I%
       '
       Close #NX%
       Call CIERRARCH("!CARMULTI")
       '
       VTB% = VENTABU%("CARMULTI")
       If VTB% >= 0 Then
         '
         NX% = FILEOPEN("FLEXMUL.DRV", 0, 79)
         Dim FMY As String * 79
         '
         For U& = 1 To ULTREG&("!CARMULTI")
           Z$ = REGLEIDO$("!CARMULTI", U&)
           RASO$ = AJUSTI$(Left$(Z$, 30), 30)
           PACCE$ = AJUSTI$(Mid$(Z$, 31, 46), 46)
           If TRIM$(RASO$) <> "" Then
             If TRIM$(PACCE$) <> "" Then
               Z1$ = String$(79, 0)
               Call REPLA(Z1$, ENCRIPTA$(RASO$), 1, 30)
               Call REPLA(Z1$, ENCRIPTA$(PACCE$), 32, 46)
               SUCAR% = 0
               For J% = 1 To 77
                 SUCAR% = SUCAR% + Asc(Mid$(Z1$, J%, 1))
               Next J%
               Call REPLA(Z1$, MKI$(SUCAR%), 78, 2)
               V& = V& + 1
               LSet FMY$ = Z1$
               Put #NX%, V&, FMY$
             End If
           End If
         Next U&
         '
         For U& = V& + 1 To LOF(NX%) / 79
           LSet FMY$ = String$(79, 0)
           Put #NX%, U&, FMY$
         Next U&
         '
         Close #NX%
         Call CIERRARCH("!CARMULTI")
         '
       End If
       '
     End If
     Call FUERASYS
     '
   End If
   '
End Sub
'
Sub CONNDOS()
   N1% = FreeFile
   Open LUBAS$ + "CODERRO.DRV" For Output As #N1%
   Print #N1%, ""
   Close #N1%
   N1% = 0
   '
   On Error Resume Next
   MkDir "\FLEXOFT"
   MkDir "\FLEXOFT\LC" + IDENTER$
   On Error GoTo 0
End Sub

