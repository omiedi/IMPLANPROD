VERSION 5.00
Begin VB.Form Form4 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Utilitarios FLEXOFT"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10065
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   10065
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      Height          =   675
      Left            =   4095
      TabIndex        =   36
      Top             =   3675
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.CommandButton Command30 
      Caption         =   "ManTec"
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
      Left            =   9030
      Picture         =   "FLEXOFT4.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Configuración General de Funcionamiento"
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
      Left            =   9030
      Picture         =   "FLEXOFT4.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   32
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
      Left            =   4620
      TabIndex        =   20
      Top             =   3990
      Width           =   4215
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
         Left            =   1470
         TabIndex        =   29
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
         Left            =   210
         TabIndex        =   28
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
         Left            =   1470
         TabIndex        =   27
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
         Left            =   2730
         TabIndex        =   26
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
         Left            =   1470
         TabIndex        =   25
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
         Left            =   210
         TabIndex        =   24
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
         Left            =   210
         TabIndex        =   23
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
         Left            =   2730
         TabIndex        =   22
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
         Left            =   2730
         TabIndex        =   21
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
      Left            =   9030
      Picture         =   "FLEXOFT4.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   19
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
      Left            =   9030
      Picture         =   "FLEXOFT4.frx":0896
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1260
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Opciones Generales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3690
      Left            =   210
      TabIndex        =   2
      Top             =   210
      Width           =   4215
      Begin VB.CommandButton Command13 
         Caption         =   "Salvapantallas"
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
         Left            =   1050
         TabIndex        =   8
         Top             =   3150
         Width           =   2010
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Colores de Pantalla"
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
         Left            =   1050
         TabIndex        =   7
         Top             =   2625
         Width           =   2010
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Clave Personal"
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
         Left            =   1050
         TabIndex        =   6
         Top             =   2100
         Width           =   2010
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Terminal"
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
         Left            =   1050
         TabIndex        =   5
         Top             =   1575
         Width           =   2010
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Monitor"
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
         Left            =   1050
         TabIndex        =   4
         Top             =   1050
         Width           =   2010
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Impresoras"
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
         Left            =   1050
         TabIndex        =   3
         Top             =   525
         Width           =   2010
      End
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
      Height          =   3690
      Left            =   4620
      TabIndex        =   1
      Top             =   210
      Width           =   4215
      Begin VB.CommandButton Command15 
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
         Height          =   330
         Left            =   1050
         TabIndex        =   14
         Top             =   3150
         Width           =   2010
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Parametros"
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
         Left            =   1050
         TabIndex        =   13
         Top             =   2625
         Width           =   2010
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
         Left            =   1050
         TabIndex        =   12
         Top             =   2100
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
         Left            =   1050
         TabIndex        =   11
         Top             =   1575
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
         Left            =   1050
         TabIndex        =   10
         Top             =   1050
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
         Left            =   1050
         TabIndex        =   9
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
      Height          =   2115
      Left            =   210
      TabIndex        =   0
      Top             =   3990
      Width           =   4215
      Begin VB.CommandButton Command28 
         Caption         =   "Parámetros de Facturación"
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
         Left            =   630
         TabIndex        =   31
         Top             =   1575
         Width           =   2955
      End
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
         Left            =   630
         TabIndex        =   30
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
         Left            =   2310
         TabIndex        =   17
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
         Left            =   630
         TabIndex        =   16
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
         Left            =   2310
         TabIndex        =   15
         Top             =   1050
         Width           =   1275
      End
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Label2"
      Height          =   225
      Left            =   2520
      TabIndex        =   34
      Top             =   0
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Label1"
      Height          =   225
      Left            =   735
      TabIndex        =   33
      Top             =   0
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   8480
      Picture         =   "FLEXOFT4.frx":0BA0
      Top             =   5580
      Width           =   2010
   End
End
Attribute VB_Name = "Form4"
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
    OPX% = COMALTER%("Elija la Función de Auditoría:\\Utilizacion\Mensajes Emergentes\Bloqueos Activos\Histórico de Bloqueos\Bloqueos Largos")
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
       ElseIf OPX% = 5 Then
         GoSub BLOQLAR
    End If
    '
    Exit Sub
    '
    '
AUDITO:
   'Call COPYSTRU("MONIEXEC", "MONIEXEQ")
   On Error Resume Next
   Kill LUDAT$ + "MONIEXEQ.DAT"
   On Error GoTo 0
   Call SELECHO("SELFECHA")
   FESEL$ = TRIM$(VALACT1$("SELFECHA"))
   If FESEL$ <> "" Then
     LU$ = LUDAT$
     N1% = FILEOPEN(LU$ + "MONIEXEC.DAT", 0, 128)
     N2% = FILEOPEN(LU$ + "MONIEXEQ.DAT", 0, 128)
     Dim XX1 As String * 128
     FIN& = LOF(N1%) / 128
     JJ& = 0
     For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       Get #N1%, U&, XX1$
       If Mid$(XX1$, 69, 8) = FESEL$ Then
         JJ& = JJ& + 1
         Put #N2%, JJ&, XX1$
       End If
     Next U&
     Close #N1%
     Close #N2%
     '
     If EXISTE%("C:\FLEXOFT\WORDPAD.EXE") = 1 Then
         Shell "C:\FLEXOFT\WORDPAD.EXE " + LUDAT$ + "MONIEXEQ.DAT", vbMaximizedFocus
       Else
         MENSA$ = "Imposible Mostrar Auditoría de Uso."
         MENSA$ = MENSA$ + "\  \Falta Utilitario 'WORDPAD.EXE'"
         MENSA$ = MENSA$ + "\en Directorio 'C:/FLEXOFT'."
         MENSA$ = MENSA$ + "\   \Instale este Programa Utilitario\y Repita la Consulta."
         Call COMUNI(MENSA$)
     End If
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
   FINI& = LOF(NX1%) / 128
   CAREMEN% = 0
   For U22& = FINI& To 1 Step -1
     Get #NX1%, U22&, TSX1$
     FXYZ% = FECHANUM(Mid$(TSX1$, 25, 8))
     If FXYZ% <= DES% Then CAREMEN% = CAREMEN% + 1
     If CAREMEN% >= 1024 Then GoTo 45
   Next U22&
   U22& = 1
   '
45 For U& = U22& To LOF(NX1%) / 128
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
BLOQLAR:
   HOO% = HOY(HO$)
   VDEF$ = MKI$(HOO%) + MKI$(HOO%)
   OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
   If OBX$ = "" Then Exit Sub
   DES% = CVI(Left$(OBX$, 2))
   HAS% = CVI(Mid$(OBX$, 3, 2))
   If DES% > HAS% Then Exit Sub
   '
   FEX = DES%
   DES1% = DIANTE(DIANTE(FEX))
   TLIM = Val(InputBox$("Filtra Bloqueos mayores que (segundos)"))
   '
   Label1.Visible = True
   Label2.Visible = True
   Screen.MousePointer = 11
   On Error Resume Next
   Kill LUDAT$ + "SEBLOFIL.DAT"
   On Error GoTo 0
   Dim TSX2 As String * 128
   JJ& = 0
   NX1% = FreeFile
   Open LUDAT$ + "SECUBLOQ.DAT" For Random Access Read Write As #NX1% Len = 128
   NX2% = FreeFile
   Open LUDAT$ + "SEBLOFIL.DAT" For Random Access Read Write As #NX2% Len = 128
   FIN& = LOF(NX1%) / 128: KKK% = 0
   For U& = FIN& To 1 Step -1
   'For U& = 1 To FIN&
     '
     If Timer > TIMAX + 0.1 Then
       TIMAX = Timer
       Label2 = JJ&: Label2.Refresh
       Label1 = Str$(FIN& - U&): Label1.Refresh
     End If
     '
     Get #NX1%, U&, TSX2$
     FXYZ% = FECHANUM(Mid$(TSX2$, 25, 8))
     If FXYZ% < DES1% Then KKK% = KKK% + 1
     If KKK% > 1000 Then GoTo 90
     If FXYZ% >= DES% Then
       If FXYZ% <= HAS% Then
         INIBLO$ = Mid$(TSX2$, 35, 8): T11& = 3600 * Val(Mid$(INIBLO$, 1, 2)) + 60 * Val(Mid$(INIBLO$, 4, 2)) + Val(Mid$(INIBLO$, 7, 2))
         FINBLO$ = Mid$(TSX2$, 55, 8): T22& = 3600 * Val(Mid$(FINBLO$, 1, 2)) + 60 * Val(Mid$(FINBLO$, 4, 2)) + Val(Mid$(FINBLO$, 7, 2))
         If Abs(T11& - T22&) > TLIM Then
           JJ& = JJ& + 1
           Put #NX2%, JJ&, TSX2$
         End If
       End If
     End If
   Next U&
90 Close #NX1%: Close #NX2%
   Label1.Visible = False
   Label2.Visible = False
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

Private Sub Command15_Click()
   Command15.Enabled = False
   If USNBR% <> 101 Then
       Call MENSERR(24, "La Opción Requiere\Privilegios de Supervisor.")
       GoTo 99
     Else
       If COMALTER%("Esta Seguro de Compactar las Bases de Datos ?\\No, Abandonar\Si, Compactar las B.Datos") = 2 Then
         Screen.MousePointer = 11
         File1.Path = LUDAT$ + "BDACCESS"
         File1.Pattern = "*.MDB"
         For U& = 1 To File1.ListCount
           NOMECOR$ = UCase$(File1.List(U& - 1))
           PPXX% = InStr(NOMECOR$, ".MDB")
           If PPXX% >= 2 Then
             NOMECOR$ = Left$(NOMECOR$, PPXX% - 1)
             On Error Resume Next
             Kill LUDAT$ + "BDACCESS\" + NOMECOR$ + "1.MDB"
             On Error GoTo 0
60           On Error Resume Next
             CompactDatabase LUDAT$ + "BDACCESS\" + NOMECOR$ + ".MDB", LUDAT$ + "BDACCESS\" + NOMECOR$ + "1.MDB"
             If Err Then GoTo 75
             FileCopy LUDAT$ + "BDACCESS\" + NOMECOR$ + ".MDB", LUDAT$ + "BDACCESS\" + NOMECOR$ + ".OLD"
             If Err Then GoTo 75
             FileCopy LUDAT$ + "BDACCESS\" + NOMECOR$ + "1.MDB", LUDAT$ + "BDACCESS\" + NOMECOR$ + ".MDB"
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
             Kill LUDAT$ + "BDACCESS\" + NOMECOR$ + "1.MDB"
             On Error GoTo 0
           End If
         Next U&
         Screen.MousePointer = 1
         Call COMUNI("Se Ha Completado la Compactacion\de las Bases de Datos '.MDB'")
       End If
   End If
99 Command15.Enabled = True
End Sub

Private Sub Command17_Click()
    Form4.Hide
End Sub

Private Sub Command11_Click()
    Call CONECRUN("#FLEXRCV", "Copias de Seguridad")
End Sub

Private Sub Command12_Click()
    Call OPNOIN("")
End Sub

Private Sub Command13_Click()
    Call OPNOIN("")
End Sub

Private Sub Command14_Click()
    Call CONECRUN("#PARAFAC", "Parámetros de Facturación")
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
    Call CONECRUN("#PARAPLA", "Definición de Ficheros Planos")
End Sub

Private Sub Command19_Click()
    Call CONECRUN("#PARALIS", "Definicion de Listadores Tabulados")
End Sub

Private Sub Command2_Click()
    Call CONECRUN("#FLEXLPT", "Seleccion de Impresoras")
End Sub

Private Sub Command20_Click()
    Call CONECRUN("#PARASTR", "Definición Estructura de Archivos")
End Sub

Private Sub Command21_Click()
    Call CONECRUN("#PARAFIL", "Procedimientos de Filtrado")
End Sub

Private Sub Command22_Click()
    Call CONECRUN("#PARAMED", "Definición de Formularios Tabulados")
End Sub

Private Sub Command23_Click()
    If EXISTE%("OBJEDIT.PRF") < 1 Then
        Call EXTRACT("OBJEDIT.PRF")
    End If
    Call CONECRUN("#PARAOBJ", "Definicion de Objetos Planos")
End Sub

Private Sub Command24_Click()
    Call CONECRUN("#PARASOR", "Definicion de Modulo de Ordenamiento")
End Sub

Private Sub Command25_Click()
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

Private Sub Command28_Click()
    Call CARDATCENCALC
End Sub

Private Sub Command29_Click()
    FLEXSETUP.Show 1
End Sub

Private Sub Command3_Click()
    Call OPNOIN("")
End Sub

Private Sub Command30_Click()
   Command30.Enabled = False
   If USNBR% Mod 100 <> 1 Then
      Call COMUNI("Requiere Privilegio de Supervisor")
      GoTo 99
   End If
   '
   OPX% = ALTERNA%("Entrar en Mantenimiento\Liberar Mantenimiento")
   If OPX% = 1 Then
       Call GRACONTROL("SYSTEM", "STATUS", "MANTENIM")
       AQUIMANTEN% = 1
     ElseIf OPX% = 2 Then
       Call GRACONTROL("SYSTEM", "STATUS", "")
   End If
99 Command30.Enabled = True
End Sub

Private Sub Command4_Click()
    Call OPNOIN("")
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
          ident$ = Mid$(F0$, 31, 12): If Left$(ident$, 1) = "@" Then ident$ = DESENCRI$(Mid$(F0$, 32, 11)) + " "
          DENO$ = Left$(F0$, 30): If Left$(DENO$, 1) = "@" Then DENO$ = DESENCRI$(Mid$(F0$, 2, 29)) + " "
          CLAV$ = Mid$(F0$, 43, 12): If Left$(CLAV$, 1) = "@" Then CLAV$ = DESENCRI$(Mid$(F0$, 44, 11)) + " "
          USU$ = TRIM$(ident$) + " - " + TRIM$(DENO$)
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
    Call CARMULTIEM
End Sub

Private Sub Command8_Click()
    Call CONECRUN("#FLEXMNU", "Configuracion de Menúes")
End Sub

Private Sub Command9_Click()
    '
    On ALTERNA%("Formularios de Impresion\Facturación - Versión 9.3\Facturación - Versión D.O.S. 97\Remitos Versión D.O.S.") GoTo 10, 20, 30, 40
    Exit Sub
10  If EXISTE%("FLFORMS.EXE") > 0 Then
        Call CONECRUN("FLFORMS", "Formularios de Impresion")
      Else
        Call CONECRUN("PRIFORMS", "Formularios de Impresion")
    End If
    Exit Sub
    '
20  Call CONECRUN("#FLEXFRM", "Diseño de Formularios")
    Exit Sub
    '
30  Call CONECRUN("#SVD171", "Formulario de Facturación")
    Exit Sub
    '
40  Call CONECRUN("#SVD172", "Formulario de Remitos")
    Exit Sub
    '
End Sub

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
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


