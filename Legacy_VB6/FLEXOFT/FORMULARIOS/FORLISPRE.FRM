VERSION 5.00
Begin VB.Form FORLISPRE 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Precios de Venta - Administración de Listas de Precios"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10065
   Icon            =   "FORLISPRE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   10065
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command12 
      Caption         =   "Llave"
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
      Left            =   9030
      Picture         =   "FORLISPRE.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   1845
      Width           =   855
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   9040
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   36
      Top             =   2700
      Width           =   855
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   37
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.ListBox SORTLIST 
      Height          =   255
      Left            =   9030
      Sorted          =   -1  'True
      TabIndex        =   35
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Complementos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   4620
      TabIndex        =   20
      Top             =   210
      Width           =   4215
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Ordenamiento y Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1960
         Left            =   315
         TabIndex        =   28
         Top             =   420
         Width           =   3585
         Begin VB.PictureBox Picture8 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   840
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   39
            Top             =   1155
            Width           =   2325
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
               Height          =   600
               Left            =   -20
               Picture         =   "FORLISPRE.frx":0454
               Style           =   1  'Graphical
               TabIndex        =   40
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "Eliminación       de Códigos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   840
               TabIndex        =   41
               Top             =   105
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture6 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   630
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   29
            Top             =   735
            Width           =   2325
            Begin VB.CommandButton Command7 
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
               Left            =   -20
               Picture         =   "FORLISPRE.frx":075E
               Style           =   1  'Graphical
               TabIndex        =   30
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Control de Códigos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   840
               TabIndex        =   31
               Top             =   0
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture7 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   420
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   32
            Top             =   315
            Width           =   2325
            Begin VB.CommandButton Command10 
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
               Left            =   -20
               Picture         =   "FORLISPRE.frx":0A68
               Style           =   1  'Graphical
               TabIndex        =   33
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label8 
               BackStyle       =   0  'Transparent
               Caption         =   "Ordenamiento"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   735
               TabIndex        =   34
               Top             =   105
               Width           =   1590
            End
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Archivos Complementarios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1510
         Left            =   315
         TabIndex        =   21
         Top             =   2490
         Width           =   3585
         Begin VB.PictureBox Picture1 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   735
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   22
            Top             =   735
            Width           =   2325
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
               Height          =   600
               Left            =   -20
               Picture         =   "FORLISPRE.frx":0D72
               Style           =   1  'Graphical
               TabIndex        =   23
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Datos Adicionales de Facturacion"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   630
               TabIndex        =   24
               Top             =   105
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture2 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   525
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   25
            Top             =   315
            Width           =   2325
            Begin VB.CommandButton Command6 
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
               Left            =   0
               Picture         =   "FORLISPRE.frx":11B4
               Style           =   1  'Graphical
               TabIndex        =   26
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Ordenes de Com- pra Terminales"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   735
               TabIndex        =   27
               Top             =   0
               Width           =   1590
            End
         End
      End
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Articulos"
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
      Left            =   9030
      Picture         =   "FORLISPRE.frx":14BE
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Alta de Nuevo Item de Stock"
      Top             =   3675
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Listas de Precios de Venta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   210
      TabIndex        =   4
      Top             =   210
      Width           =   4215
      Begin VB.Frame Frame8 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Actualizacion y Consulta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   315
         TabIndex        =   6
         Top             =   2310
         Width           =   3585
         Begin VB.PictureBox Picture4 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   735
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   10
            Top             =   840
            Width           =   2325
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
               Left            =   -20
               Picture         =   "FORLISPRE.frx":1608
               Style           =   1  'Graphical
               TabIndex        =   11
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "Emisión Lista por Impresora"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   735
               TabIndex        =   12
               Top             =   105
               Width           =   1380
            End
         End
         Begin VB.PictureBox Picture3 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   525
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   7
            Top             =   420
            Width           =   2325
            Begin VB.CommandButton Command3 
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
               Left            =   0
               Picture         =   "FORLISPRE.frx":1C72
               Style           =   1  'Graphical
               TabIndex        =   8
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Actualizacion de Precios"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   735
               TabIndex        =   9
               Top             =   0
               Width           =   1275
            End
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Configuración"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   315
         TabIndex        =   5
         Top             =   420
         Width           =   3585
         Begin VB.PictureBox Picture11 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   735
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   14
            Top             =   840
            Width           =   2325
            Begin VB.CommandButton Command9 
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
               Left            =   0
               Picture         =   "FORLISPRE.frx":1F7C
               Style           =   1  'Graphical
               TabIndex        =   15
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "Lista de Pre-   cios Existente"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   735
               TabIndex        =   16
               Top             =   105
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture5 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   525
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   17
            Top             =   420
            Width           =   2325
            Begin VB.CommandButton Command5 
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
               Left            =   0
               Picture         =   "FORLISPRE.frx":2286
               Style           =   1  'Graphical
               TabIndex        =   18
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "Nueva Lista"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   735
               TabIndex        =   19
               Top             =   105
               Width           =   1590
            End
         End
      End
   End
   Begin VB.CommandButton Command1 
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
      Height          =   750
      Left            =   9030
      Picture         =   "FORLISPRE.frx":2590
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Terminar - Salir de la Aplicación"
      Top             =   315
      Width           =   855
   End
   Begin VB.CommandButton Command29 
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
      Height          =   750
      Left            =   9030
      Picture         =   "FORLISPRE.frx":26DA
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Ayuda Flexoft en Línea"
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Direct"
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
      Left            =   9030
      Picture         =   "FORLISPRE.frx":29E4
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Acceso Directo a Otros Procesos"
      Top             =   2910
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "    Sistema Modular Integrado de Procesamiento de Datos                              Linea FLEXOFT (r) - Version 9.03"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   3780
      TabIndex        =   0
      Top             =   4830
      Width           =   6315
   End
   Begin VB.Image Image3 
      Height          =   885
      Left            =   0
      Picture         =   "FORLISPRE.frx":2CEE
      Stretch         =   -1  'True
      Top             =   4620
      Width           =   10920
   End
End
Attribute VB_Name = "FORLISPRE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
   'WindowState = 0
   'Refresh
End Sub

Private Sub Form_Load()
   '
   Call VERJOBID
   '
   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
      Caption = Caption + "  -  " + EPAC$
   End If
   Call GRATITFOR(Caption)
   '
   Show
   Call APRECIO1
   '
End Sub
'
Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     OPENFORMS = DoEvents
   Wend
End Sub

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Private Sub Command12_Click()
    Command12.Enabled = False
5   CLAX$ = ""
10  PPX% = Int(Rnd * 63)
    If PPX% < 48 Or PPX% > 57 Then GoTo 10
    CLAX$ = CLAX$ + Chr$(PPX%)
    If Len(CLAX$) < 6 Then GoTo 10
    '
    RESPUES$ = ""
    For M% = 1 To 6
        POSI% = (5 * M%) Mod 6
        While POSI% < 1: POSI% = POSI% + 6: Wend
        VORI% = Val(Mid$(CLAX$, POSI%, 1))
        VTRA% = (VORI% + 3 * M%) Mod 10
        RESPUES$ = RESPUES$ + Mid$(Str$(VTRA%), 2)
    Next M%
    '
    VDEF$ = CLAX$ + Space$(6)
    CTRS$ = OBJPLA$("CLACONTRA/Generar Llave de Acceso", VDEF$)
    If CTRS$ = "" Then
        Call FINAL("")
    End If
    CONTRASE$ = Mid$(CTRS$, 7, 6)
    If CONTRASE$ <> RESPUES$ Then
        INTEN% = INTEN% + 1
        If INTEN% < 3 Then GoTo 5
        Call FINAL("")
    End If
    '
15  Call COMUNI("Coloque Diskette en Blanco")
    SVL$ = SERIVOL$("A:")
    If SVL$ = "" Then GoTo 15
    TECONTROL$ = TRIM$(UCase$(App.EXEName)) + SVL$
    NX1% = FreeFile
    Dim TTX As String * 128
    On Error GoTo 90
    Open "A:\" + TRIM$(UCase$(App.EXEName)) + ".KEY" For Random Access Read Write As #NX1% Len = 128
    Randomize Timer
    TTY$ = ""
    For kki% = 1 To 128
      TTY$ = TTY$ + Chr$(Int(255 * Rnd))
    Next kki%
    LTC% = Len(TECONTROL$)
    Call REPLA(TTY$, MKI$(LTC%), 18, 2)
    Call REPLA(TTY$, ENCRIPTA$(TECONTROL$), 30, LTC%)
    TTX$ = TTY$
    Put #NX1%, 2, TTX$
    Close #NX1%
    On Error GoTo 0
    Call COMUNI("Llave Generada")
    GoTo 99
    '
90  Resume 91
91  On Error GoTo 0
    Call COMUNI("Imposible Generar LLave")
    '
99  Command12.Enabled = True
End Sub

Private Sub Command10_Click()
    Command10.Enabled = False
    Call APRECIO6
    Command10.Enabled = True
End Sub

Private Sub Command11_Click()
        '
        Command11.Enabled = False
        '
1000    ' ******************************** eliminacion de codigos
        '
        X$ = REGSEL$("LISTAS")
        If Len(X$) <= 4 Then
            GoTo 99
        End If
        '
1010    X$ = OBJPLA$("LISPRE-MUESTRA", X$)
        If X$ = "" Then
            GoTo 99
        End If
        CADECI% = Asc(Mid$(X$, 49, 1))
        If CADECI% < 2 Or CADECI% > 4 Then CADECI% = 2
        '
        LISNO% = Asc(Left$(X$, 1))
        ALISTA$ = TRIM$(Str$(LISNO%)): While Len(ALISTA$) < 3
                                    ALISTA$ = "0" + ALISTA$
                                  Wend
        Call DATLIPRE(X$)        ' datos de lista elegida
        Call CIERRARCH("*.*")
        '
        LIPREAC$ = "LIPRE" + ALISTA$
        VDEF$ = MKI$(0) + MKI$(0) + Chr$(0)
        '
10      Screen.MousePointer = 1
        OBX$ = OBJPLA$("DESHASCOD", VDEF$)
        If OBX$ = "" Then GoTo 99
        '
        Screen.MousePointer = 11
        CI1% = CVI(Left$(OBX$, 2))
        If CI1% < 1 Or CI1% > NUMAS% Then GoTo 10
        COD1$ = CODEXT$(CI1%)
        CI2% = CVI(Mid$(OBX$, 3, 2))
        If CI1% < 1 Or CI1% > NUMAS% Then GoTo 10
        coD2$ = CODEXT$(CI2%)
        If coD2$ < COD1$ Then GoTo 10
        If TRIM$(COD1$) = "" Then GoTo 10
        If TRIM$(coD2$) = "" Then GoTo 10
        '
        Screen.MousePointer = 1
        If COMALTER%("Confirme la Eliminación del Rango de Códigos\desde " + TRIM$(COD1$) + "  hasta " + TRIM$(coD2$) + "\\Cancelar\Confirmar") <> 2 Then
          GoTo 10
        End If
        '
        Screen.MousePointer = 11
        Call COPYSTRU("LIPREVEN", LIPREAC$)
        JJ& = 0
        FIN& = ULTREG&(LIPREAC$)
        For U& = 1 To FIN&
          Call TRACE(20, U&, FIN&)
          X$ = REGLEIDO$(LIPREAC$, U&)
          CIXI% = CVI(Left$(X$, 2))
          If CIXI% > 0 Then
            If CIXI% <= NUMAS% Then
              CDX$ = CODEXT$(CIXI%)
              If CDX$ < COD1$ Or CDX$ > coD2$ Then
                JJ& = JJ& + 1
                Call GRAREG(LIPREAC$, X$, JJ&)
              End If
            End If
          End If
        Next U&
        '
        Call SETULTREG(LIPREAC$, JJ&)
        Call CIERRARCH(LIPREAC$)
        Screen.MousePointer = 1
        '
99 Command11.Enabled = True
   '
End Sub

Private Sub Command19_Click()
    Call CONECRUN("AMASTO01", "Maestro de Items de Stock")
    Call FRESHECHO("MASTER")
End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    Call FINAL("*DATADI")
    Command2.Enabled = True
End Sub

Private Sub Command3_Click()
    '
    Command3.Enabled = False
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    '
    If InStr(EPAC$, "MAINIERI") > 0 Or InStr(EPAC$, "SEALU") > 0 Or InStr(EPAC$, "DURALAM") > 0 Or InStr(EPAC$, "FASS") > 0 Then
       OPX% = COMALTER%("Puede Actualizar Precios en Forma Interactiva\o por Porcentaje sobre el Precio Base\\Interactiva\% s/Precio Base\Cancelar")
       If OPX% = 1 Then
            Call APRECIO2
          ElseIf OPX% = 2 Then
            If DERACCE%("APREPOR", "Actualización de Precios por Porcentaje") = 2 Then
              Call ACOSPOR
            End If
       End If
       GoTo 99
    End If
    '
    OPX% = COMALTER%("Puede Actualizar Precios en Forma MANUAL DIRECTA\o por PORCENTAJE sobre el Precio Vigente\\Manual Directa\% s/Precio Vigente\Cancelar")
    If OPX% = 1 Then
         Call APRECIO2
       ElseIf OPX% = 2 Then
         If DERACCE%("APREPOR", "Actualización de Precios por Porcentaje") = 2 Then
            Call APREPOR
         End If
    End If
    '
99  Command3.Enabled = True
    '
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    Call APRECIO3
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    Call APRECIO4
    Command5.Enabled = True
End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
    Call FINAL("*OCOTERM")
    Command6.Enabled = True
End Sub

Private Sub Command7_Click()
    Command7.Enabled = False
    Call APRECIO7
    Command7.Enabled = True
End Sub

Private Sub Command8_Click()
    Call ACCDIR("ADMIPRE")
End Sub

Private Sub Command9_Click()
    Command9.Enabled = False
    Call APRECIO5
    Command9.Enabled = True
End Sub

Sub ACOSPOR()
    '
    Dim VALVIE(32767), VALNUE(32767)
    '
    PORCALPRE = Val(TRIM$(CONTROL$("", "PORCALPRE")))
    If PORCALPRE < 1 Then
      Call MENSERR(24, "Proceso Imposible.\Falta Definir Porcentaje de Cálculo de Precios")
      Exit Sub
    End If
    '
    VDEF$ = ""
10  OBX$ = OBJPLA$("APREBAPOR", VDEF$)
    If OBX$ = "" Then Exit Sub
    '
    VDEF$ = OBX$
    CI1% = CVI(Left$(OBX$, 2)): COD1$ = CODEXT$(CI1%)
    CI2% = CVI(Mid$(OBX$, 3, 2)): coD2$ = CODEXT$(CI2%)
    '
    If CI1% < 1 Or CI2% < 1 Or TRIM$(COD1$) = "" Or TRIM$(coD2$) = "" Then
       Call MENSERR(24, "Códigos no Válidos")
       GoTo 10
    End If
    '
    If coD2$ < COD1$ Then
       Call MENSERR(24, "Códigos no Válidos")
       GoTo 10
    End If
    '
    PORAJU = CVS(Mid$(OBX$, 5, 4))
    If PORAJU < (-90) Or PORAJU > 99.99 Then
       Call MENSERR(24, "Porcentaje no Válido")
       GoTo 10
    End If
    '
    CADECI% = CVI(Mid$(OBX$, 9, 2))
    If CADECI% < 2 Or CADECI% > 4 Then
       Call MENSERR(24, "Decimales entre '2' y '4'")
       GoTo 10
    End If
    '
    Screen.MousePointer = 11
    JJ& = 0: FIN& = NUINV%
    REBLA$ = REGBLAN$("ACOSPOR")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      I0$ = REGLEIDO$("INVERT", U& + 1)
      I1$ = TRIM$(Left$(I0$, 16))
      If Len(I1$) <= 15 Then
        If I1$ >= COD1$ Then
          If I1$ <= coD2$ Then
            CIXI% = CVI(Mid$(I0$, 17, 2))
            If CIXI% > 0 Then
              If CIXI% <= NUMAS% Then
                REMA& = CIXI% + 1
                REMAS$ = REGLEIDO$("MASTER", REMA&)
                COSVIE = CVS(Mid$(REMAS$, 111, 4))
                COSNUE = COSVIE * (100 + PORAJU) / 100
                COSNUE = (Int(COSNUE * 10 ^ CADECI% + 0.5)) / (10 ^ CADECI%)
                ZZ$ = REBLA$
                Call REPLA(ZZ$, MKI$(CIXI%), 1, 2)
                Call REPLA(ZZ$, MKS$(COSVIE), 3, 4)
                Call REPLA(ZZ$, MKS$(COSNUE), 7, 4)
                Call REPLA(ZZ$, MKS$(REMA&), 11, 4)
                JJ& = JJ& + 1
                Call GRAREG("!ACOSPOR", ZZ$, JJ&)
              End If
            End If
          End If
        End If
      End If
    Next U&
    '
    Call SETULTREG("!ACOSPOR", JJ&)
    Call CIERRARCH("!ACOSPOR")
    Screen.MousePointer = 1
    '
    VTB% = VENTABU%("ACOSPOR")
    If VTB% >= 0 Then
      If COMALTER%("Actualiza Costos y Precios de Venta ?\\Cancelar\Actualizar") = 2 Then
         '
         URE& = ULTREG&("!ACOSPOR")
         For U& = 1 To URE&
           Call TRACE(20, U&, URE&)
           X$ = REGLEIDO$("!ACOSPOR", U&)
           CIXI% = CVI(Left$(X$, 2))
           If CIXI% > 0 Then
             If CIXI% <= NUMAS% Then
               REMA& = CIXI% + 1
               REMAS$ = REGLEIDO$("MASTER", REMA&)
               If CVI(Left$(REMAS$, 2)) = CIXI% Then
                 If Abs((CVS(Mid$(REMAS$, 111, 4))) - (CVS(Mid$(X$, 3, 4)))) < 0.005 Then
                   VALVIE(CIXI%) = CVS(Mid$(X$, 3, 4))
                   VALNUE(CIXI%) = CVS(Mid$(X$, 7, 4))
                   Call REPLA(REMAS$, MKS$(VALNUE(CIXI%)), 111, 4)
                   Call GRAREG("MASTER", REMAS$, REMA&)
                 End If
               End If
             End If
           End If
         Next U&
         Call CIERRARCH("MASTER")
         '
         FIN& = ULTREG&("LIPRE001")
         For U& = 1 To FIN&
           Call TRACE(20, U&, FIN&)
           X$ = REGLEIDO$("LIPRE001", U&)
           CIXI% = CVI(Left$(X$, 2))
           If CIXI% > 0 Then
             If CIXI% <= NUMAS% Then
               PRENUE# = CDbl((Int((100 + PORCALPRE) * VALNUE(CIXI%) + 0.5)) / 100)
               If PRENUE# > 0.005 Then
                 Call REPLA(X$, MKD$(PRENUE#), 62, 8)
                 Call GRAREG("LIPRE001", X$, U&)
               End If
             End If
           End If
         Next U&
         Call CIERRARCH("LIPRE001")
         '
         Call COMUNI("Actualización Completa !!!")
         '
      End If
    End If
    '
End Sub

Sub APREPOR()
    '
    Dim VALVIE(32767), VALNUE(32767)
    HOII% = HOY(HOS$)
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "DOSIVA") > 0 Then
      If VERILLAVE% <> 1 Then Exit Sub
    End If
    '
    X$ = REGSEL$("LISTAS")
    If Len(X$) <= 4 Then
      Exit Sub
    End If
    '
    RANGOAC% = 0
    If InStr(EPAC$, "DOSIVA") > 0 Then
      TTXX$ = "Opciones de Actualización por Porcentaje:\\Sólo Bombas\Sólo Repuestos\Toda la Lista"
      RANGOAC% = COMALTER%(TTXX$)
      If RANGOAC% < 1 Or RANGOAC% > 3 Then Exit Sub
      If RANGOAC% = 3 Then RANGOAC% = 0
    End If
    '
    Select Case RANGOAC%
      Case 0
        RACTU$ = "Lista Completa"
      Case 1
        RACTU$ = "Bombas Nacionales (Cod. 0...)"
      Case 2
        RACTU$ = "Repuestos (Cod. >= 1...)"
    End Select
    '
    VDEF$ = Left$(X$, 64) + MKS$(0)
    Call REPLA(VDEF$, MKI$(HOII%), 63, 2)
    CADECI% = Asc(Mid$(VDEF$, 49, 1))
    If CADECI% < 2 Or CADECI% > 4 Then CADECI% = 2
    Call REPLA(VDEF$, Chr$(CADECI%), 49, 1)
    '
5   OBX$ = OBJPLA$("APREVENPOR/Actualizacion por Porcentaje - " + RACTU$, VDEF$)
    If OBX$ = "" Then
      Exit Sub
    End If
    '
    PORAJU = CVS(Mid$(OBX$, 65, 4))
    If Abs(PORAJU) < 0.001 Then Exit Sub
    '
    If PORAJU < (-90) Or PORAJU > 99.99 Then
       Call MENSERR(24, "Porcentaje no Válido")
       GoTo 5
    End If
    '
    CADECI% = Asc(Mid$(OBX$, 49, 1))
    If CADECI% < 2 Or CADECI% > 4 Then
       Call MENSERR(24, "Decimales entre '2' y '4'")
       GoTo 5
    End If
    '
    LISNO% = Asc(Left$(OBX$, 1))
    ALISTA$ = TRIM$(Str$(LISNO%)): While Len(ALISTA$) < 3
                                     ALISTA$ = "0" + ALISTA$
                                   Wend
                                   '
    APRECIO$ = "LIPRE" + ALISTA$
    Call COPYSTRU("LIPREVEN", APRECIO$)
    REBLA$ = REGBLAN$("ACOSPOR")
    FIN& = ULTREG&(APRECIO$)
    JJ& = 0
    '
    Screen.MousePointer = 11
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      I0$ = REGLEIDO$(APRECIO$, U&)
      CIXI% = CVI(Left$(I0$, 2))
      '
      PORAJX = PORAJU
      If RANGOAC% > 0 Then
        CDDX$ = Left$(CODEXT$(CIXI%), 1)
        If RANGOAC% = 1 Then
            If CDDX$ <> "0" Then PORAJX = 0
          ElseIf RANGOAC% = 2 Then
            If CDDX$ = "0" Then PORAJX = 0
        End If
      End If
      '
      COSVIE = CVD(Mid$(I0$, 62, 8))
      COSNUEX$ = CSTRING$(MKS$(COSVIE * (100 + PORAJX) / 100), 3, 12, CADECI%, 2)
      COSNUE = Val(COSNUEX$)
      ZZ$ = REBLA$
      Call REPLA(ZZ$, MKI$(CIXI%), 1, 2)
      Call REPLA(ZZ$, MKS$(COSVIE), 3, 4)
      Call REPLA(ZZ$, MKS$(COSNUE), 7, 4)
      JJ& = JJ& + 1
      Call GRAREG("!ACOSPOR", ZZ$, JJ&)
    Next U&
    '
    Call SETULTREG("!ACOSPOR", JJ&)
    Call CIERRARCH("!ACOSPOR")
    Screen.MousePointer = 1
    '
    VTB% = VENTABU%("ACOSPOR/TITUPAN=Actualizacion de Precios por Porcentaje - " + RACTU$)
    If VTB% >= 0 Then
      If COMALTER%("Desea Guardar los Nuevos Precios de Venta ?\\Cancelar\Guardar") = 2 Then
         '
         Screen.MousePointer = 11
         URE& = ULTREG&("!ACOSPOR")
         For U& = 1 To URE&
           Call TRACE(20, U&, URE&)
           X$ = REGLEIDO$("!ACOSPOR", U&)
           Y$ = REGLEIDO$(APRECIO$, U&)
           MRX$ = Mid$(Y$, 45, 1)
           COSNUEX$ = CSTRING$(Mid$(X$, 7, 4), 3, 12, CADECI%, 2)
           PRENUEVO# = Val(COSNUEX$)
           PREX$ = COSNUEX$
           Call REPLA(Y$, MKD$(PRENUEVO#), 62, 8)
           Call REPLA(Y$, PREX$ + "  ", 43 - Len(PREX$), 2 + Len(PREX$))
           Call REPLA(Y$, MRX$, 44, 1)
           Call GRAREG(APRECIO$, Y$, U)
         Next U&
         '
         Call CIERRARCH(APRECIO$)
         Call BAJALDISCO
         Screen.MousePointer = 1
         '
         Call COMUNI("Actualización Completa !!!")
         '
      End If
    End If
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub
