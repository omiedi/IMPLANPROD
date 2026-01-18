VERSION 5.00
Begin VB.Form FSetupSql 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Config. FLEXOFT"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6255
   Icon            =   "FSETUPSQL.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   6255
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Licen."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Left            =   5250
      TabIndex        =   33
      ToolTipText     =   "Licenciamiento Aplicaciones Complementarias"
      Top             =   4740
      Width           =   855
   End
   Begin VB.CommandButton Command30 
      Caption         =   "Dis.Frm"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Left            =   5250
      TabIndex        =   32
      ToolTipText     =   "Diseñador de Formularios"
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton Command12 
      Caption         =   "SQL-Srvr"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Left            =   5250
      TabIndex        =   24
      Top             =   4140
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Acc-Dir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Left            =   5250
      TabIndex        =   23
      Top             =   3840
      Width           =   855
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   5250
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   30
      Top             =   3600
      Width           =   855
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   31
         Top             =   105
         Width           =   12000
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Compacta Cta.Corr."
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
      Left            =   5250
      TabIndex        =   29
      Top             =   5145
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.FileListBox File1 
      Height          =   675
      Left            =   5640
      Pattern         =   "*.MDB"
      TabIndex        =   28
      Top             =   4830
      Visible         =   0   'False
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
      Height          =   750
      Left            =   5250
      Picture         =   "FSETUPSQL.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   2560
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
      Height          =   750
      Left            =   5250
      Picture         =   "FSETUPSQL.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Configuración General de Funcionamiento"
      Top             =   1810
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
      TabIndex        =   27
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   13
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
         TabIndex        =   19
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   14
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
         TabIndex        =   16
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
      Height          =   750
      Left            =   5250
      Picture         =   "FSETUPSQL.frx":0B8E
      Style           =   1  'Graphical
      TabIndex        =   0
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
      Height          =   750
      Left            =   5250
      Picture         =   "FSETUPSQL.frx":0CD8
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1060
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
      TabIndex        =   26
      Top             =   1890
      Width           =   4845
      Begin VB.CommandButton Command14 
         Caption         =   "Parametros de Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1780
         TabIndex        =   6
         Top             =   420
         Width           =   1290
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Roll-Back"
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
         Left            =   3255
         TabIndex        =   10
         ToolTipText     =   "Restaura Version Anterior de Ejecutable o Controlador"
         Top             =   1050
         Width           =   1290
      End
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
         Left            =   1780
         TabIndex        =   9
         Top             =   1050
         Width           =   1290
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Compactar B.Datos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   3240
         TabIndex        =   7
         Top             =   420
         Width           =   1290
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
         Left            =   315
         TabIndex        =   8
         Top             =   1050
         Width           =   1290
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
         Height          =   450
         Left            =   315
         TabIndex        =   5
         Top             =   420
         Width           =   1290
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
      TabIndex        =   25
      Top             =   225
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
         TabIndex        =   1
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
         TabIndex        =   2
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
         TabIndex        =   3
         ToolTipText     =   "Restauración de B.Datos Desde Copia de Seguridad"
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
         Left            =   2745
         TabIndex        =   4
         Top             =   1050
         Width           =   1275
      End
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   4695
      Picture         =   "FSETUPSQL.frx":0FE2
      Top             =   5145
      Width           =   2010
   End
End
Attribute VB_Name = "FSetupSql"
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
    OPX% = COMALTER%("Elija la Función de Auditoría:\\Utilizacion\Mensajes Emergentes\Bloqueos Activos\Histórico de Bloqueos\Busqueda por Texto")
    If OPX% = 1 Then
         GoSub AUDITO
       ElseIf OPX% = 2 Then
      
         Call SELECHO("SELFECHA")
         FFXX$ = VALACT1$("SELFECHA")
         If FFXX$ = "" Then Exit Sub
         FFII% = FECHANUM(FFXX$)
         JJ& = 0
         Screen.MousePointer = 11
         JJ& = 0
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
       ElseIf OPX% = 5 Then
         CADENA$ = ""
         CADENA$ = InputBox("Texto de Busqueda:", "Ingresar Texto")
         If TRIM(CADENA$) = "" Then
            Exit Sub
         End If
         Screen.MousePointer = 11
         JJ& = 0
         For U& = 1 To ULTREG&("MENSAJES")
            XX$ = REGLEIDO$("MENSAJES", U&)
            If InStr(1, UCase(XX$), UCase(CADENA$)) > 1 Then
               JJ& = JJ& + 1
               Call GRAREG("MENSAFIL", XX$, JJ&)
            End If
         Next U&
         Call SETULTREG("MENSAFIL", JJ&)
         Call CIERRARCH("MENSAFIL")
         Call CIERRARCH("MENSAJES")
         Screen.MousePointer = 1
         Call CONECRUN("*LIMENSA", "Listado de Mensajes Emergentes")
    End If
    '
    Exit Sub
    '
    '
AUDITO:
   RUWOPA$ = "C:\FLEXOFT\"
   If EXISTE%(RUWOPA + "WORDPAD.EXE") < 1 Then
     RUWOPA$ = ""
     If EXISTE%(RUWOPA + "WORDPAD.EXE") < 1 Then
       MENSA$ = "Imposible Mostrar Auditoría de Uso."
       MENSA$ = MENSA$ + "\  \Falta Utilitario 'WORDPAD.EXE'"
       MENSA$ = MENSA$ + "\en Directorio 'C:/FLEXOFT'."
       MENSA$ = MENSA$ + "\   \Instale este Programa Utilitario\y Repita la Consulta."
       Call COMUNI(MENSA$)
       GoTo 90
     End If
   End If
   '
20 MODOI% = ALTERNA%("Secuencia por Fecha\Procesos Interrumpidos\Bases de Datos Recuperadas")
   If MODOI% = 1 Then
       Call SELECHO("SELFECHA")
       FFXX$ = VALACT1$("SELFECHA")
       If FFXX$ = "" Then Exit Sub
       Screen.MousePointer = 11
       NX1% = FILEOPEN%(LUDAT$ + "MONIEXEC.DAT", 0, 128)
       NX2% = FILEOPEN%(LUCOM$ + "MONIEXEK.DAT", 2, 128)
       CANRE& = LOF(NX1%) / 128
       On Error Resume Next
       Dim AX1 As String * 128
       On Error GoTo 0
       For U& = 1 To CANRE&
         Call TRACE(20, U&, CANRE&)
         Get #NX1%, U&, AX1$
         If Mid$(AX1$, 69, 8) = FFXX$ Then
            Print #NX2%, AX1$;
         End If
       Next U&
       Close #NX2%
       Close #NX1%
       Screen.MousePointer = 1
       Shell RUWOPA$ + "WORDPAD.EXE " + LUCOM$ + "MONIEXEK.DAT", vbMaximizedFocus
       '
     ElseIf MODOI% = 2 Then
     Call DESHASFECHA(DES%, HAS%, PERIO$) ' RE- IMPUTACION
     If PERIO$ = "" Then
      GoTo 20
     End If

     DESDEX = FECHATEX(DES%)
     HASTAX = FETEXCOR$(HAS%)

     
'       Call SELECHO("SELFECHA")
'       FECX$ = VALACT1$("SELFECHA")
'       If FECX$ = "" Then Exit Sub
       Screen.MousePointer = 11

       JJ& = 0
       Dim BB As String * 128
       NX1% = FILEOPEN(LUDAT$ + "MONIEXEC.DAT", 0, 128)
       NX2% = FILEOPEN(LUDAT$ + "MONIBRK.PRN", 0, 128)
       FIN& = LOF(NX1%) / 128
       '
       For U& = 1 To FIN&
         Call TRACE(20, U&, FIN&)
         Get #NX1%, U&, BB$
         fe% = FECHANUM(Mid$(BB$, 69, 8))
         If fe% >= DES% And fe% <= HAS% Then
        'If Mid$(BB$, 69, 8) >= DESDEX And Mid$(BB$, 69, 8) <= HASTAX Then
            If TRIM$(Mid$(BB$, 89, 18)) = "" Then
              JJ& = JJ& + 1
              Put #NX2%, JJ&, BB$
            End If
        End If
       Next U&
       '
       Screen.MousePointer = 1
       Close #NX1%
       Close #NX2%
       Shell RUWOPA$ + "WORDPAD.EXE " + LUDAT$ + "MONIBRK.PRN", vbMaximizedFocus
     ElseIf MODOI% = 3 Then
       Shell RUWOPA$ + "WORDPAD.EXE " + LUDAT$ + "RECUBASE.PRN", vbMaximizedFocus
   End If

90 Return
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

Private Sub Command12_Click()
    CONN_SQLS.Show 1
End Sub

Private Sub Command13_Click()
    Command13.Enabled = False
    If DERACCE%("ROLLBACK", "Restaurar Versiones Anteriores") >= 2 Then
       SELECARCH.Show 1
    End If
99 Command13.Enabled = True
End Sub

Private Sub Command14_Click()
    Command14.Enabled = False
    PARACON.Show 1
    Command14.Enabled = True
End Sub

Private Sub Command17_Click()
    Call CIERRARCH("*.*")
    Call COMUNI("Re-Configuración Completa.\Re-Inicie Ahora la Aplicación.")
    Call FINAL("")
End Sub

Private Sub Command11_Click()
    Command11.Enabled = False
    If DERACCE%("RECUBACK", "Restauración de B.Datos desde Backup") < 2 Then GoTo 99
    '
    CUDI$ = App.Path
    LUBA$ = LUBAS$
    HOII% = HOY(HOS$)
    '
    LUDA$ = LUDAT$
    If Mid$(LUDA$, 2, 1) = ":" Then LUDA$ = Mid$(LUDA$, 3)
    If Left$(LUDA$, 2) = "\\" Then LUDA$ = Mid$(LUDA$, 2)
    '
    On Error Resume Next
    MkDir ULOWP$ + "BACKUPS"
    LUDACO$ = LUDA$
    PPX1% = 1
    While Len(LUDACO$) > 0
      PPX1% = InStr(1 + PPX1%, LUDACO$, "\")
      If PPX1% < 1 Then GoTo 15
      MkDir ULOWP$ + "BACKUPS" + Left$(LUDACO$, PPX1% - 1)
    Wend
15  On Error GoTo 0
    '
    AA1 = RECLEN("INDIAPLI")
    Call COPYSTRU("INDIAPLI", "INDIBAK")
    JJ& = 0
    RUTABAK$ = "C:\FLEXOFT\BACKUPS" + LUDA$
    EXISZ$ = TRIM$(Dir(RUTABAK$ + "*.ARJ"))
20  If EXISZ$ <> "" Then
      FIDA$ = Format$(FileDateTime(RUTABAK$ + EXISZ$), "dd/mm/yy - hh:mm")
      JJ& = JJ& + 1
      TTXX$ = AJUSTI$(EXISZ$, 12) + "Generado el " + FIDA$ + " hs"
      Call GRAREG("INDIBAK", TTXX$, JJ&)
      EXISZ$ = TRIM$(Dir)
      GoTo 20
    End If
    Call SETULTREG("INDIBAK", JJ&)
    Call FRESHECHO("INDIBAK")
    '
    If JJ& < 1 Then
      Call MENSERR(24, "No Existen Respaldos Recuperables\en esta Terminal.\  \Pruebe en Otras Estaciones de Trabajo.")
      GoTo 99
    End If
    '
    Call SELECHO("INDIBAK")
    ASELE$ = TRIM$(VALACT1$("INDIBAK"))
    If ASELE$ = "" Then GoTo 99
    '
    TTXX$ = "Al Restaurar Datos de un Respaldo Anterior\Se Produce la Pérdida Irrecuperable\de los Datos Actuales.\  \Esta Transacción no Podrá Revertirse !!!\  \Está Seguro de Continuar Restaurando los Datos\desde el Archivo '" + ASELE$ + "' " + VALACT2$("INDIBAK") + " ?"
    If COMALTER%(TTXX$ + "\\No, Cancelar\Sí, Continuar") <> 2 Then GoTo 99
    '
    '
30  On Error Resume Next
    FileCopy CurDir + "\ARJ.EXE", LUDAT$ + "ARJ.EXE"
    On Error GoTo 0
    '
    CMD$ = ""
    UTIBAK$ = "ARJ E -Y"
    ABAK$ = RUTABAK$ + ASELE$
    LUDAX$ = LUDAT$
    If Right$(LUDAX$, 1) = "\" Then LUDAX$ = Left$(LUDAX$, Len(LUDAX$) - 1)
    If Mid$(LUDAX$, 2, 1) = ":" Then
      CMD$ = CMD$ + Left$(LUDAX$, 2) + CRLF$
      LUDAX$ = Mid$(LUDAX$, 3)
    End If
    '
    CMD$ = CMD$ + "CHDIR " + LUDAX$ + " " + CRLF$
    CMD$ = CMD$ + UTIBAK$ + " " + ABAK$ + CRLF$
    '
    If Mid$(CUDI$, 2, 1) = ":" Then
      CMD$ = CMD$ + Left$(CUDI$, 2) + CRLF$
      CUDI$ = Mid$(CUDI$, 3)
    End If
    CMD$ = CMD$ + "CHDIR " + CUDI$ + CRLF$
    CMD$ = CMD$ + "ECHO OFF" + CRLF$
    CMD$ = CMD$ + "ECHO Fin de la Operacion de Restauración" + CRLF$
    CMD$ = CMD$ + "PAUSE " + CRLF$
    CMD$ = CMD$ + "EXIT " + CRLF$
    Call SAVEFILE(LUCOM$ + "DOSCOMM.BAT", CMD$)
    Call SAVEFILE(CurDir + "\RECUBAK", "RECUBAK")
    '
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    '
    MODOBA = vbNormalFocus
    SBX1 = Shell(LUCOM$ + "DOSCOMM.BAT", MODOBA)
    '
99  Command11.Enabled = True
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
20  MODO% = ALTERNA%("A-B-M Usuarios del Sistema\Derechos de Acceso por Usuario\Copia de Derechos de Acceso entre Usuarios")
    If MODO% = 1 Then
          Call ACTUSER
        ElseIf MODO% = 2 Then
          Call ACDERACC
        ElseIf MODO% = 3 Then
          Call CODERACC
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
   '
''''   Dim BACUECORR As Database
''''   Set BACUECORR = OpenDatabase(LUACC$ + "CUECORRI.MDB")
''''   '
''''   Set TABLORIGI = BACUECORR.TableDefs("CAJABANC")
''''   '
''''   ' crea la base de datos
''''   Screen.MousePointer = 11
''''   YACREADA% = 0
''''   For KKI% = 1 To TABLORIGI.Fields.Count
''''     NOMCAMPO$ = TABLORIGI.Fields(KKI% - 1).Name
''''     TICAM% = TABLORIGI.Fields(KKI% - 1).Type
''''     LOCAM% = TABLORIGI.Fields(KKI% - 1).Size
''''     On Error Resume Next
''''     If YACREADA% = 0 Then
''''       Set CreaTabla = BACUECORR.CreateTableDef("CABANVIE")
''''       CreaTabla.Fields.Append CreaTabla.CreateField(NOMCAMPO$, TICAM%, LOCAM%)
''''       BACUECORR.TableDefs.Append CreaTabla
''''       Set CREACAMPOS = BACUECORR.TableDefs("CABANVIE")
''''       YACREADA% = 1
''''     End If
''''     '
''''     With CREACAMPOS.Fields
''''       .Append CREACAMPOS.CreateField(NOMCAMPO$, TICAM%, LOCAM%)
''''       If TICAM% = 1 Or TICAM% = 10 Or TICAM% = 12 Then
''''         CREACAMPOS.Fields(NOMCAMPO$).AllowZeroLength = True
''''       End If
''''     End With
''''     On Error GoTo 0
''''     '
''''   Next KKI%
''''   '
''''   FELIMITE = CDbl(CVDAT(FECHANUM("01/07/05")))
''''   Set TABLDESTI = BACUECORR.TableDefs("CABANVIE")
''''   '
''''   Call ABRECABAN
''''   Screen.MousePointer = 11
''''   Set TABLDESTI = CueCorri.OpenRecordset("CABANVIE", dbOpenDynaset)
''''   CajaBanc.MoveLast
''''   FIN& = CajaBanc.RecordCount
''''   CajaBanc.MoveFirst
''''   Do While Not CajaBanc.EOF
''''     If CDbl(CajaBanc!FechConta) < FELIMITE Then
''''       TABLDESTI.AddNew
''''       For KKI% = 1 To CajaBanc.Fields.Count
''''         NOMCAMPO$ = CajaBanc.Fields(KKI% - 1).Name
''''         TABLDESTI.Fields(NOMCAMPO$).Value = CajaBanc.Fields(NOMCAMPO$).Value
''''       Next KKI%
''''       TABLDESTI.Update
''''       CajaBanc.Delete
''''     End If
''''     KKL& = KKL& + 1
''''     Call TRACE(20, KKL&, FIN&)
''''   CajaBanc.MoveNext
''''   Loop
''''   '
''''   Call ABRECABAN
''''   Screen.MousePointer = 11
''''   SQLX$ = "SELECT * FROM CABANVIE "
''''   SQLX$ = SQLX$ + "WHERE NuClien > 0 "
''''   SQLX$ = SQLX$ + "AND (TipoMovim = 'FVEN' "
''''   SQLX$ = SQLX$ + "OR TipoMovim = 'RCOB' "
''''   SQLX$ = SQLX$ + "OR TipoMovim = 'AJUD') "
''''   SQLX$ = SQLX$ + "ORDER BY FechEmisi, fehoreal ASC;"
''''   Set CuecoTemp = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
''''   '
''''   JJ& = 0: KU& = 0 ' SALDO# = 0: SARRAS# = 0
''''   REBLA$ = REGBLAN$("CUECOCLI")
''''   With CuecoTemp
''''     On Error Resume Next
''''     .MoveLast
''''     If Err < 1 Then
''''       On Error GoTo 0
''''       FIN& = .RecordCount
''''       .MoveFirst
''''       Do While Not .EOF
''''         KU& = KU& + 1
''''         Call TRACE(20, KU&, FIN&)
''''         TIX$ = Left$(!CodiCom_Lar, 2)
''''         SG% = (-1)
''''         TIPOS$ = "Factura": If TIX$ = "AS" Then TIPOS$ = "Ajuste de Saldo"
''''                             If TIX$ = "AJ" Then TIPOS$ = "Ajuste de Saldo"
''''                             If TIX$ = "RB" Then TIPOS$ = "Recibo Cobranza"
''''                             If TIX$ = "VC" Then TIPOS$ = "Venta Contado"
''''                             If TIX$ = "ND" Then TIPOS$ = "Nota de Debito"
''''                             If TIX$ = "DC" Then TIPOS$ = "Descuento Caja"
''''                             If TIX$ = "NC" Then TIPOS$ = "Nota de Credito"
''''         IMPO# = !IDEBECOMP - !IHABECOMP
''''         '
''''         FEXI% = CVINT%(!FechEmisi)
''''         NROL& = !NumeComp
''''         NCLIE% = !NuClien
''''         '
''''         Z$ = REBLA$
''''         Call REPLA(Z$, MKI$(NCLIE%), 1, 2)
''''         Call REPLA(Z$, MKI$(FEXI%), 3, 2)
''''         Call REPLA(Z$, TIPOS$, 5, 16)
''''         Call REPLA(Z$, MKS$(NROL&), 21, 4)
''''         Call REPLA(Z$, MKD$(IMPO#), 25, 8)
''''         Call REPLA(Z$, !CodiCom_Lar, 33, 24)
''''         Call REPLA(Z$, !CodiCom_Cor, 57, 12)
''''         Call REPLA(Z$, MKI$(!CodiEmpr), 69, 2)
''''         Call REPLA(Z$, MKI$(!CodiSucu), 71, 2)
''''         Call REPLA(Z$, MKI$(!Status), 73, 2)
''''         JJ& = JJ& + 1
''''         Call GRAREG("CUECOCLI", Z$, JJ&)
''''         '
''''       .MoveNext
''''       Loop
''''       Call SETULTREG("CUECOCLI", JJ&)
''''     End If
''''   End With
''''   '
''''   Screen.MousePointer = 11
''''   SQLX$ = "SELECT * FROM CABANVIE "
''''   SQLX$ = SQLX$ + "WHERE NuProve > 0 "
''''   SQLX$ = SQLX$ + "AND (TipoMovim = 'FCOM' "
''''   SQLX$ = SQLX$ + "OR TipoMovim = 'OPAG' "
''''   SQLX$ = SQLX$ + "OR TipoMovim = 'AJUP') "
''''   SQLX$ = SQLX$ + "ORDER BY FechEmisi ASC, FeHoReal ASC;"
''''   Set CuecoTemp = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
''''   '
''''   JJ& = 0: KU& = 0 ' SALDO# = 0: SARRAS# = 0
''''   REBLA$ = REGBLAN$("CUECOPRO")
''''   With CuecoTemp
''''     On Error Resume Next
''''     .MoveLast
''''     If Err < 1 Then
''''       On Error GoTo 0
''''       FIN& = .RecordCount
''''       .MoveFirst
''''       Do While Not .EOF
''''         KU& = KU& + 1
''''         Call TRACE(20, KU&, FIN&)
''''         TIX$ = Left$(!CodiCom_Lar, 2)
''''         SG% = (-1)
''''         TIPOS$ = "Factura": If TIX$ = "AS" Then TIPOS$ = "Aj.Saldo"
''''                             If TIX$ = "TF" Then TIPOS$ = "Tkt-Factura"
''''                             If TIX$ = "RH" Then TIPOS$ = "Rec.Honor."
''''                             If TIX$ = "ND" Then TIPOS$ = "N.Debito"
''''                             If TIX$ = "CP" Then TIPOS$ = "Créd.Propio"
''''                             If TIX$ = "CO" Then TIPOS$ = "Comprobante"
''''                             If TIX$ = "OP" Then TIPOS$ = "O.Pago": SG% = 1
''''                             If TIX$ = "FF" Then TIPOS$ = "O/P F.Fijo": SG% = 1
''''                             If TIX$ = "NC" Then TIPOS$ = "N.Crédito": SG% = 1
''''                             If TIX$ = "DP" Then TIPOS$ = "Deb.Propio": SG% = 1
''''                             '
''''         IMPO# = !IDEBECOMP - !IHABECOMP
''''         '
''''         FEXI% = CVINT%(!FechEmisi)
''''         NROX$ = !CodiCom_Lar
''''         NROL& = !NumeComp
''''         NCLIE% = !NuProve
''''         '
''''         Z$ = REBLA$
''''         Call REPLA(Z$, MKI$(NCLIE%), 1, 2)
''''         Call REPLA(Z$, MKI$(FEXI%), 3, 2)
''''         Call REPLA(Z$, TIPOS$, 5, 16)
''''         Call REPLA(Z$, MKS$(NROL&), 21, 4)
''''         Call REPLA(Z$, MKD$(IMPO#), 25, 8)
''''         Call REPLA(Z$, NROX$, 33, 24)
''''         Call REPLA(Z$, !CodiCom_Cor, 57, 12)
''''         Call REPLA(Z$, MKI$(!CodiEmpr), 69, 2)
''''         Call REPLA(Z$, MKI$(!CodiSucu), 71, 2)
''''         Call REPLA(Z$, MKI$(!Status), 73, 2)
''''         JJ& = JJ& + 1
''''         Call GRAREG("CUECOPRO", Z$, JJ&)
''''       .MoveNext
''''       Loop
''''       Call SETULTREG("CUECOPRO", JJ&)
''''     End If
''''   End With
''''   '
''''   Screen.MousePointer = 1
   '
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

Private Sub Command23_Click()
    If EXISTE%("OBJEDIT.PRF") < 1 Then
        Call EXTRACT("OBJEDIT.PRF")
    End If
    Call PARAOBJ
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

Private Sub Command3_Click()
    '\\\OT-07-0002-RG-03/01/07///
    Command3.Enabled = False
    If CLACONTRA% > 0 Then
       On Error Resume Next
       HABILAP.Show 1
       If Err Then
          On Error GoTo 0
          Call MENSERR(24, "Falta Formulario 'HABILAP.FRM' en Proyecto '" + App.EXEName + "'.")
       End If
       On Error GoTo 0
    End If
    Command3.Enabled = True
    '\\\OT-07-0002-RG-FIN///
End Sub

Private Sub Command30_Click()
     Command30.Enabled = False
     Call CONECRUN("DSGFRM09", "Diseñador de Formularios")
     Command30.Enabled = True
End Sub

Private Sub Command4_Click()
   Call ABRETACON
   SQLX$ = "SELECT * FROM SQL_ACCDIR "
   Set CONSULTEMP = BaDaConsul.OpenRecordset(SQLX$, dbOpenDynaset)
   U& = 0
   Call BLANARCH("TACONSUL")
   Y$ = REGBLAN$("TACONSUL")
   With CONSULTEMP
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then
        On Error GoTo 0
        Do While .EOF = False
            U& = U& + 1
            EJECU$ = !Ejecutable
            LEYEN$ = !LEYENDA
            CODDD$ = !Codigo
            '
            'CARGA EL ARCHIVO CON LOS DATOS DE LA TABLA
            '
            Call REPLA(Y$, CODDD$, 1, 16)
            Call REPLA(Y$, LEYEN$, 16, 64)
            Call REPLA(Y$, EJECU$, 80, 16)
            Call GRAREG("!TACONSUL", Y$, U&)
            .MoveNext
        Loop
     End If
     Call SETULTREG("!TACONSUL", U&)
   End With
   VTB% = VENTABU%("TACONSUL")
   '
   SQLX$ = "SELECT * FROM  SQL_ACCDIR"
   Set MARCATEMP = BaDaConsul.OpenRecordset(SQLX$, dbOpenDynaset)
   On Error Resume Next
   MARCATEMP.MoveFirst
   If Err < 1 Then
     Do While Not MARCATEMP.EOF
      MARCATEMP.Edit
      MARCATEMP!Status = 1
      MARCATEMP.Update
      MARCATEMP.MoveNext
     Loop
   End If
   
   
   SQLX$ = "SELECT * FROM  SQL_ACCDIR"
   Set CONSUTEMP = BaDaConsul.OpenRecordset(SQLX$, dbOpenDynaset)
   With CONSUTEMP
   FIN& = ULTREG&("!TACONSUL")
     For I& = 1 To FIN&
         TXX1$ = REGLEIDO$("!TACONSUL", I&)
         CODDX$ = TRIM$(Mid$(TXX1$, 1, 16))
         LEYEX$ = TRIM$(Mid$(TXX1$, 16, 64))
         EJEXU$ = TRIM$(Mid$(TXX1$, 80, 16))
         '
         .AddNew
         !Ejecutable = EJEXU$
         !LEYENDA = LEYEX$
         !Codigo = CODDX$
         .Update
     Next I&
   End With
   
   SQLX$ = "DELETE * FROM SQL_ACCDIR "
   SQLX$ = SQLX$ + "WHERE Status =" & 1
   BaDaConsul.Execute (SQLX$)
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
          CLAV$ = TRIM$(CLAV$)
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
    DIRETRAB04.Show
    'Call DIRETRA
End Sub

Private Sub Command7_Click()
   Command7.Enabled = False
   Call CARMULTIEM
   Command7.Enabled = True
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   If USNBR% <> 101 Then
       Call MENSERR(24, "Este Proceso Requiere\Privilegios de Supervisor.")
     Else
       If COMALTER%("Esta Seguro de Compactar las Bases de Datos ?\\No, Abandonar\Si, Compactar las B.Datos") = 2 Then
         Screen.MousePointer = 11
         rutax$ = LUDAT$ + "BDACCESS"
20       File1.Path = rutax$
         File1.Pattern = "*.MDB"
         For U& = 1 To File1.ListCount
           NOMECOR$ = UCase$(File1.List(U& - 1))
           PPXX% = InStr(NOMECOR$, ".MDB")
           If PPXX% >= 2 Then
             NOMECOR$ = Left$(NOMECOR$, PPXX% - 1)
             On Error Resume Next
             Kill rutax$ + "\" + NOMECOR$ + "1.MDB"
             On Error GoTo 0
60           On Error Resume Next
             CompactDatabase rutax$ + "\" + NOMECOR$ + ".MDB", rutax$ + "\" + NOMECOR$ + "1.MDB"
             If Err Then GoTo 75
             FileCopy rutax$ + "\" + NOMECOR$ + ".MDB", rutax$ + "\" + NOMECOR$ + ".OLD"
             If Err Then GoTo 75
             FileCopy rutax$ + "\" + NOMECOR$ + "1.MDB", rutax$ + "\" + NOMECOR$ + ".MDB"
             If rutax$ = App.Path Then
                FileCopy rutax$ + "\" + NOMECOR$ + "1.MDB", rutax$ + "\" + NOMECOR$ + ".KMP"
             End If
             If Err Then
75             TTXX$ = "Error (" & Err.Number & ") al Compactar " + NOMECOR$
               If Err.Number = 3343 Then
                  On Error GoTo 0
                  TTXX$ = TTXX$ + "\  \Esta Base de Datos está Dañada."
                  TTXX$ = TTXX$ + "\¿ Desea Que FLEXOFT Intente Repararla ?"
                  TTXX$ = TTXX$ + "\\Si, Reparar\No, Cancelar"
                  If COMALTER%(TTXX$) = 1 Then
                    On Error Resume Next
                    DBEngine.RepairDatabase LUACC$ + NOMECOR$ + ".MDB"
                    GoTo 60
                  End If
                  GoTo 80
               End If
               MsgBox TTXX$
             End If
80           On Error GoTo 0
             On Error Resume Next
             Kill rutax$ + "\" + NOMECOR$ + "1.MDB"
             On Error GoTo 0
           End If
         Next U&
         '
         If rutax$ = LUDAT$ + "BDACCESS" Then
           If rutax$ <> App.Path Then
             rutax$ = App.Path
             GoTo 20
           End If
         End If
         '
         Screen.MousePointer = 1
         Call COMUNI("Se Ha Completado la Compactacion\de las Bases de Datos '.MDB'")
       End If
   End If
   Command8.Enabled = True
End Sub

Private Sub Command9_Click()
    '
    Command9.Enabled = False
    ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
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
    Call ABRECONEXION
    '
    If APLINVO$ <> "" Then
       On Error Resume Next
       If APLINVO$ = "CLAVUSER" Then
           Call Command5_Click                ' CAMBIO CLAVE PERSONAL DE ACCESO
         ElseIf APLINVO$ = "ANACOSTO" Then
           OpAnCoAcc.Show 1
         ElseIf APLINVO$ = "OPCLAVEN" Then
           OPCLAV04.Show 1
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
           OPTRAJU04.Show 1
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
Sub PARAOBJ()
     '
     GoSub 7000
     '
10   MODO% = ALTERNA%("Nuevo Objeto\Objeto Existente\-\Mostrar Objeto\-\Terminar")
     On MODO% GoTo 100, 20, 10, 900
     GoTo 3900

20   Copia% = 0
     X$ = REGSEL$("OBJNAME")
     If Len(X$) <= 4 Then GoTo 10
     X$ = REGLEIDO$("OBJNAME", CVS(Right$(X$, 4)))
     GoSub 6000                                          ' leer datos
     GoTo 150

100  Copia% = 0
     OPXX% = COMALTER%("¿ Copia de Objeto Existente ?\\Si, Copiar\No, Directo")
     If OPXX% = 1 Then
         X$ = REGSEL$("OBJNAME")
         If Len(X$) <= 4 Then GoTo 10
         X$ = REGLEIDO$("OBJNAME", CVS(Right$(X$, 4)))
         GoSub 6000
         OBJPUNT& = 0
         OBJORI$ = Left$(X$, 16)
         Copia% = 1
         GoTo 105
       ElseIf OPXX% = 2 Then
         X$ = REGBLAN$("OBJNAME")
         Call REPLA(X$, MKI$(2), 65, 2)
         Call REPLA(X$, String$(190, 0), 67, 190)
105      OBJNAME$ = UCase$(TRIM$(InputBox$("Codigo", "Identificador del Objeto")))
         If OBJNAME$ = "" Then GoTo 10
         Call REPLA(X$, AJUSTI$(OBJNAME$, 16), 1, 16)
         GoTo 110
       Else
         GoTo 10
     End If
     '
110  OBJNAME1$ = AJUSTI$(OBJNAME$, 16)
     RFF$ = RECFILT$("OBJNAME", 1, OBJNAME1$)
     If RFF$ <> "" Then
        Call MENSERR(24, "Ya Existe un Objeto\con ese Código.")
        GoTo 10
     End If
     '
150    VDEF$ = X$
160    OBX$ = OBJPLA$("ENC-OBJETO", VDEF$)
       If OBX$ = "" Then GoTo 10
       '
       OBJNAME$ = AJUSTI$(OBJNAME$, 16)
       RFF$ = RECFILT$("OBJDETA", 1, OBJNAME$)
       OBJPUNT& = CVS(RFF$ + String$(4, 0))
       Call REPLA(OBX$, MKS$(OBJPUNT&), 67, 4)
       Call FILTERPUTRECORD("OBJNAME", 1, OBJNAME$, OBX$)
       '
300    ORI$ = OBJNAME$
       If Copia% = 1 Then ORI$ = OBJORI$
       '
       Call FILTERFILE("OBJDETA", "", 1, ORI$)
       VTB% = VENTABU%("OBJDETA")
       If VTB% >= 0 Then
          For KU& = 1 To ULTREG&("!OBJDETA")
            XXZ$ = REGLEIDO$("!OBJDETA", KU&)
            Call REPLA(XXZ$, OBJNAME$, 1, 16)
            
            Call GRAREG("!OBJDETA", XXZ$, KU&)
          Next KU&
          Call FILTERUPDATE("OBJDETA", "!OBJDETA", 1, OBJNAME$)
          '
          RFF$ = RECFILT$("OBJDETA", 1, OBJNAME$)
          OBJPUNT& = CVS(RFF$ + String$(4, 0))
          OBX$ = FILTERGETRECORD$("OBJNAME", 1, OBJNAME$)
          Call REPLA(OBX$, MKS$(OBJPUNT&), 67, 4)
          Call FILTERPUTRECORD("OBJNAME", 1, OBJNAME$, OBX$)
          '
          GoSub 7100
          '
       End If
       '
       Call FRESHECHO("OBJNAME")
       GoTo 10

900    X$ = REGSEL$("OBJNAME")
       If Len(X$) <= 4 Then GoTo 10
       X$ = REGLEIDO$("OBJNAME", CVS(Right$(X$, 4)))
       OBX$ = Mid$(X$, 1, 16)
       XXX$ = OBJPLA$(OBX$, "")                   ' mostrar
       '
       Call FRESHECHO("OBJNAME")
       GoTo 10

3900   Close: End

6000   OBJNAME$ = Mid$(X$, 1, 16)
       OBJDESC$ = Mid$(X$, 17, 48)
       OBJTYPE1% = CVI(Mid$(X$, 65, 2))
       If OBJTYPE1% = 1 Then
           OBJTYPE$ = "1.- Caja de Mensajes     "
         ElseIf OBJTYPE1% = 2 Then
           OBJTYPE$ = "2.- Caja Entrada de Datos"
       End If
       OBJPUNT& = CVS(Mid$(X$, 67, 4))
       '
       BSUP% = Asc(Mid$(X$, 71, 1))
       BIZQ% = Asc(Mid$(X$, 72, 1))
       ALTU% = Asc(Mid$(X$, 73, 1))
       ALIB% = Asc(Mid$(X$, 74, 1))
       TITU$ = Mid$(X$, 75, 32)
       COLO$ = Mid$(X$, 107, 3)
       CIERI% = Asc(Mid$(X$, 110, 1)): CIER$ = "NO"
           If CIERI% = 1 Then CIER$ = "SI"
       CONFI% = Asc(Mid$(X$, 111, 1)): CONF$ = "NO"
           If CONFI% = 1 Then CONF$ = "SI"
           '
       Return

6100   X$ = REGACT$("OBJNAME")
       X$ = REPLACE$(X$, MKS$(OBJPUNT&), 67, 4)
       X$ = REPLACE$(X$, Chr$(BSUP%), 71, 1)
       X$ = REPLACE$(X$, Chr$(BIZQ%), 72, 1)
       X$ = REPLACE$(X$, Chr$(ALTU%), 73, 1)
       X$ = REPLACE$(X$, Chr$(ALIB%), 74, 1)
       X$ = REPLACE$(X$, TITU$, 75, 32)
       X$ = REPLACE$(X$, COLO$, 107, 3)
       X$ = REPLACE$(X$, Chr$(CIERI%), 110, 1)
       X$ = REPLACE$(X$, Chr$(CONFI%), 111, 1)
       '
       REG& = REGNUM&("OBJNAME")
       Call GRAREG("OBJNAME", X$, REG&)
       Call CIERRARCH("OBJNAME")
       '
       Return

7000   ULORI$ = Left$(CurDir$, 2)
       FDT1# = 0
       On Error Resume Next
       FDT1# = FileDateTime(ULORI$ + "\FLEXOFT\NOVEDAD\OBJNAME.DRV")
       On Error GoTo 0
       FDT2# = 0
       On Error Resume Next
       FDT2# = FileDateTime("F:\FLEXOFT\NOVEDAD\OBJNAME.DRV")
       On Error GoTo 0
       '
       On Error Resume Next
       Kill LUDAT$ + "OBJNAME.DAT"
       Kill LUDAT$ + "OBJDETA.DAT"
       On Error GoTo 0
       '
       If FDT2# >= FDT1# Then ULORI$ = "F:"
       ULORI$ = ULORI$ + "\FLEXOFT\NOVEDAD\"
       LBX$ = ULORI$
       Dim XYZ1 As String * 256
       NX1% = FILEOPEN%(ULORI$ + "OBJNAME.DRV", 0, 256)
       JJ& = 0
       For KKI& = 1 To LOF(NX1%) / 256
         Get #NX1%, KKI&, XYZ1$
         If TRIM$(Left$(XYZ1$, 16)) <> "" Then
           JJ& = JJ& + 1
           Call GRAREG("OBJNAME", XYZ1$, JJ&)
         End If
       Next KKI&
       Call SETULTREG("OBJNAME", JJ&)
       Close #NX1%: NX1% = 0
       '
       Dim XYZ2 As String * 128
       NX1% = FILEOPEN%(ULORI$ + "OBJDETA.DRV", 0, 128)
       JJ& = 0
       For KKI& = 1 To LOF(NX1%) / 128
         Get #NX1%, KKI&, XYZ2$
         If TRIM$(Left$(XYZ2$, 16)) <> "" Then
           JJ& = JJ& + 1
           Call GRAREG("OBJDETA", XYZ2$, JJ&)
         End If
       Next KKI&
       Call SETULTREG("OBJDETA", JJ&)
       Close #NX1%: NX1% = 0
       '
Return



7100   LBY$ = LUBAS$
       AA1$ = LOADFILE$(LUDAT$ + "OBJNAME.DAT")
       Call SAVEFILE(LBX$ + "OBJNAME.DRV", AA1$)
       Call SAVEFILE(LBY$ + "OBJNAME.DRV", AA1$)
       AA1$ = LOADFILE$(LUDAT$ + "OBJDETA.DAT")
       Call SAVEFILE(LBX$ + "OBJDETA.DRV", AA1$)
       Call SAVEFILE(LBY$ + "OBJDETA.DRV", AA1$)
       '
Return

End Sub


Sub CODERACC() 'Copia derechos de acceso entre usuarios
    '
    OBX$ = OBJPLA$("DEUSUAUS", VDEF$)
    If OBX$ = "" Then Exit Sub
    USORIG% = CVI(Left$(OBX$, 2))
    USDEST% = CVI(Mid$(OBX$, 3, 2))
    
    V& = 0
    For U& = 1 To ULTREG&("FLEXACC")
        X$ = REGLEIDO$("FLEXACC", U&)
        UI% = CVI(Left$(X$, 2))
        If CVI(Left$(X$, 2)) = USDEST% Then
            OPX% = COMALTER%("El Usuario 'Destino' Posee Derechos, Desea Sobreescribirlo?\\No\Si")
            If OPX% = 2 Then
                GoTo 10
            Else
                Exit Sub
            End If
        End If
    Next U&
    '
    Call ABRETACON
    SQLX$ = "SELECT * FROM DER_ACCE "
    SQLX$ = SQLX$ + " WHERE NUM_USER = " & USDEST%
    Set PARATEMP = BaDaConsul.OpenRecordset(SQLX$, dbOpenDynaset)
    With PARATEMP
        On Error Resume Next
        .MoveFirst
        If Err Then
           On Error GoTo 0
           PARATEMP.Close
           GoTo 10
        End If
    End With
    PARATEMP.Close
    '
    On Error GoTo 0
    OPX% = COMALTER%("El Usuario 'Destino' Posee Derechos, Desea Sobreescribirlo?\\No\Si")
    If OPX% = 2 Then
           GoTo 10
        Else
           Exit Sub
    End If
    '
10  On Error Resume Next
    PARATEMP.Close
    On Error GoTo 0
    Call BLOQARCH("FLEXACC")
    '
    VACON% = 1
    For U& = 1 To ULTREG&("FLEXACC")
      X$ = REGLEIDO$("FLEXACC", U&)
      If CVI(Left$(X$, 2)) <> USDEST% Then
          V& = V& + 1
          Call GRAREG("FLEXACC", X$, V&)
      End If
    Next U&
    '
    Call SETULTREG("FLEXACC", V&)
    '
    On Error Resume Next
    SQLX$ = "DELETE * FROM DER_ACCE WHERE NUM_USER =  " & USDEST%
    BaDaConsul.Execute SQLX$
    On Error GoTo 0
    '
    For U& = 1 To ULTREG&("FLEXACC")
      X$ = REGLEIDO$("FLEXACC", U&)
      If CVI(Left$(X$, 2)) = USORIG% Then
         If Mid$(X$, 11, 1) = "S" Or Mid$(X$, 12, 1) = "S" Then
            Z1$ = ""
            Call REPLA(Z1$, MKI$(USDEST%), 1, 2)
            Call REPLA$(Z1$, Mid$(X$, 3, 8), 3, 8)
            Call REPLA(Z1$, Mid$(X$, 11, 1), 11, 1)
            Call REPLA(Z1$, Mid$(X$, 12, 1), 12, 1)
            V& = V& + 1
            Call REGAPP("FLEXACC", Z1$)
         End If
      End If
    Next U&
    '
    Call SETULTREG("FLEXACC", V&)
    Call CIERRARCH("FLEXACC")
    Call LIBARCH("FLEXACC")
    '
    SQLX$ = "SELECT * FROM DER_ACCE "
    SQLX$ = SQLX$ + " WHERE NUM_USER = " & USORIG%
    Set PARATEMP = BaDaConsul.OpenRecordset(SQLX$, dbOpenSnapshot)
    With PARATEMP
       On Error Resume Next
       .MoveFirst
       On Error GoTo 0
       Do While Not .EOF
          CAPLI$ = !COD_APLI
          DRX1% = !DER_READ
          DRX2% = !DER_WRITE
          '
          SQLY$ = "SELECT * FROM DER_ACCE "
          SQLY$ = SQLY$ + " WHERE NUM_USER = " & USDEST%
          Set PARADEST = BaDaConsul.OpenRecordset(SQLY$, dbOpenDynaset)
          PARADEST.AddNew
          PARADEST!NUM_USER = USDEST%
          PARADEST!COD_APLI = CAPLI$
          PARADEST!DER_READ = DRX1%
          PARADEST!DER_WRITE = DRX2%
          PARADEST.Update
          PARADEST.Close
       .MoveNext
       Loop
    End With
    PARATEMP.Close
    '
    Call MENSERR(24, "Proceso de Copia Finalizado")
    '
End Sub

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


