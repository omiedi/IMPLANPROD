VERSION 5.00
Begin VB.Form ORCOMP00 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ordenes de Compra - Proveedores Externos"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5865
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command19 
      Caption         =   "N.Orig"
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
      Left            =   4830
      Picture         =   "ORCOMP00.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Códigos Equivalentes de Proveedor"
      Top             =   4305
      Width           =   855
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Maestro"
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
      Left            =   4830
      Picture         =   "ORCOMP00.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Maestro de Items de Stock"
      Top             =   2835
      Width           =   855
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   4830
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5145
      Width           =   855
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.CommandButton Command18 
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
      Height          =   720
      Left            =   4830
      Picture         =   "ORCOMP00.frx":0454
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Configuración de Procesos y Formularios"
      Top             =   1785
      Width           =   855
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Proveed"
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
      Left            =   4830
      Picture         =   "ORCOMP00.frx":0896
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "A-B-M Maestro de Proveedores"
      Top             =   3570
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0FFC0&
      Caption         =   "Gestión de Compras"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5790
      Left            =   210
      TabIndex        =   7
      Top             =   210
      Width           =   4425
      Begin VB.Frame Frame8 
         BackColor       =   &H00E0FFC0&
         Caption         =   "Control de Gestión"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2955
         Left            =   420
         TabIndex        =   19
         Top             =   2625
         Width           =   3585
         Begin VB.Frame Frame6 
            BackColor       =   &H00E0FFC0&
            Caption         =   "Salida"
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
            Left            =   210
            TabIndex        =   29
            Top             =   2205
            Width           =   3200
            Begin VB.OptionButton Option1 
               BackColor       =   &H00E0FFC0&
               Caption         =   "Pantalla"
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
               Left            =   420
               TabIndex        =   31
               Top             =   250
               Value           =   -1  'True
               Width           =   1380
            End
            Begin VB.OptionButton Option2 
               BackColor       =   &H00E0FFC0&
               Caption         =   "Impresión"
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
               TabIndex        =   30
               Top             =   250
               Width           =   1275
            End
         End
         Begin VB.PictureBox Picture5 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   840
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   20
            Top             =   1365
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
               Left            =   -20
               Picture         =   "ORCOMP00.frx":0BA0
               Style           =   1  'Graphical
               TabIndex        =   21
               ToolTipText     =   "Pendientes por Material"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "Por Código del Material"
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
               TabIndex        =   22
               Top             =   105
               Width           =   1200
            End
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   630
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   23
            Top             =   900
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
               Picture         =   "ORCOMP00.frx":0EAA
               Style           =   1  'Graphical
               TabIndex        =   24
               ToolTipText     =   "Pendientes por Proveedor"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "Por Proveedor"
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
               TabIndex        =   25
               Top             =   105
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture3 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   420
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   26
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
               Picture         =   "ORCOMP00.frx":12EC
               Style           =   1  'Graphical
               TabIndex        =   27
               ToolTipText     =   "Ordenes de Compra por Número"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "O/C's por Número"
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
               TabIndex        =   28
               Top             =   105
               Width           =   1590
            End
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0FFC0&
         Caption         =   "Ordenes de Compra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2010
         Left            =   420
         TabIndex        =   8
         Top             =   420
         Width           =   3585
         Begin VB.PictureBox Picture8 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   945
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   11
            Top             =   1155
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
               Left            =   1680
               Picture         =   "ORCOMP00.frx":15F6
               Style           =   1  'Graphical
               TabIndex        =   18
               ToolTipText     =   "Revertir Anulaciones de Ordenes de Compra"
               Top             =   0
               Width           =   560
            End
            Begin VB.CommandButton Command14 
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
               Picture         =   "ORCOMP00.frx":1740
               Style           =   1  'Graphical
               TabIndex        =   2
               ToolTipText     =   "Anulación de Orden de Compra"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "Anulación O/Compras"
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
               TabIndex        =   12
               Top             =   105
               Width           =   1065
            End
         End
         Begin VB.PictureBox Picture14 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   630
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   13
            Top             =   735
            Width           =   2325
            Begin VB.CommandButton Command22 
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
               Picture         =   "ORCOMP00.frx":1A4A
               Style           =   1  'Graphical
               TabIndex        =   1
               ToolTipText     =   "Ver por Pantalla y Re-Imprimir"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "Ver y Re-Imprimir"
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
               TabIndex        =   14
               Top             =   105
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture16 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   315
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   15
            Top             =   315
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
               Left            =   0
               Picture         =   "ORCOMP00.frx":1D54
               Style           =   1  'Graphical
               TabIndex        =   0
               ToolTipText     =   "Generar una Nueva Orden de Compra"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label17 
               BackStyle       =   0  'Transparent
               Caption         =   "Nueva O/Compra"
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
               TabIndex        =   16
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
      Height          =   720
      Left            =   4830
      Picture         =   "ORCOMP00.frx":205E
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
      Height          =   720
      Left            =   4830
      Picture         =   "ORCOMP00.frx":21A8
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Ayuda Flexoft en Línea"
      Top             =   1050
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   4275
      Picture         =   "ORCOMP00.frx":24B2
      Top             =   5460
      Width           =   2010
   End
End
Attribute VB_Name = "ORCOMP00"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Hide
End Sub

Private Sub Command10_Click()
   Command10.Enabled = False
   Call DESANULOC
   Command10.Enabled = True
End Sub

Private Sub Command11_Click()
    Command11.Enabled = False
    AMASTO00.Show 1
99  Command11.Enabled = True
End Sub


Private Sub Command14_Click()
   Command14.Enabled = False
   Call OCOMANU
   Command14.Enabled = True
End Sub

Private Sub Command16_Click()
    Command16.Enabled = False
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "NEUMANN") > 0 Then
        Call COCOBREC
      Else
        Call CONECRUN("AMAPRO", "Actualización de Base de Datos de Proveedores")
    End If
    Command16.Enabled = True
End Sub

Private Sub Command18_Click()
    'OPOCOMPRA.Show 1
    Call CONECRUN("FLSETUP/OPOCOMPRA", "Configuración de Funcionamiento")
    Call FINAL("")
End Sub

Private Sub Command19_Click()
   Command19.Enabled = False
   Call FINAL("*CODORPR")
   Command19.Enabled = True
End Sub

Private Sub Command22_Click()
    '
    Command22.Enabled = False
    '
3   Call SELECHO("OCOMENCA/Indice General de Ordenes de Compra (Pendientes-Cumplidas-Anuladas)")
    NPX& = Val(VALACT1$("OCOMENCA"))
    If NPX& < 1 Then GoTo 99
    '
    TIDOCUM$ = "Nota de Pedido"
    NUDOCU$ = TRIM$(Str$(NPX&))
    While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
    NUDOCU$ = " " + NUDOCU$ + " "
    Screen.MousePointer = 11
    '
    For UI& = ULTREG&("PDOCLST") To 1 Step -1
      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(XX$, TIDOCUM$) > 4 Then
        PPXX% = InStr(XX$, "Nro.")
        If PPXX% > 0 Then
          XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
        End If
        If InStr(XX$, NUDOCU$) > 4 Then
          DOCUNU& = CVS(Left$(XX$, 4))
          GoTo 5
        End If
      End If
    Next UI&
    Screen.MousePointer = 1
    Call COMUNI("Documento no Encontrado")
    GoTo 99
    '
5   For KKI% = 0 To 5
      EDITFORM.Picture1(KKI%).Cls
      EDITFORM.Picture1(KKI%).DrawWidth = 3
      EDITFORM.Picture1(KKI%).DrawStyle = 0
      EDITFORM.Picture1(KKI%).FillStyle = 1
    Next KKI%
    '
    REDOCU$ = REGACT$("PDOCLST")
    URE& = ULTREG&("PDOCSPL")
    LI& = 0: LS& = URE&
6   E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
    L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
    DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
    If DCC& < DOCUNU& Then LI& = L&: GoTo 6
    If DCC& > DOCUNU& Then LS& = L&: GoTo 6
    GoTo 8
    '
7   Screen.MousePointer = 1
    Call MENSERR(24, "Documento no Registrado")
    GoTo 99
    '
8   TPSP$ = ""
    Screen.MousePointer = 11
    While L& > 1
      DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
      If DCC& < DOCUNU& Then GoTo 9
      L& = L& - 1
    Wend
    '
9   CAPAGINAS% = 0
    For KKL& = L& To URE&
      TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
        TPSP$ = TPSP$ + Mid$(TBUF$, 5)
      End If
    Next KKL&
    '
95  SPOOLSTRING$ = TPSP$
    FIN& = Len(TPSP$)
    For KKI% = 0 To 5
      EDITFORM.Picture1(KKI%).Cls
      EDITFORM.Picture1(KKI%).Top = 0
      EDITFORM.Picture1(KKI%).Refresh
      EDITFORM.Picture1(KKI%).Height = 7070
      EDITFORM.Picture1(KKI%).Width = 11700
      EDITFORM.Picture1(KKI%).DrawWidth = 3
      EDITFORM.Picture1(KKI%).DrawStyle = 0
      EDITFORM.Picture1(KKI%).FillStyle = 1
    Next KKI%
    ALTUPAGINA = 7070
    TOPEPAGINA = 0
    PAGINACTIVA% = 0
    CAPAGINAS% = 0
    '
10  PPXX% = InStr(TPSP$, Chr$(255))
    If PPXX% > 0 Then
      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
      TPSP$ = Mid$(TPSP$, PPXX% + 1)
      Call TRACE(20, FIN& - Len(TPSP$), FIN&)
      '
      If Len(COPRIAT$) > 0 Then
        COPRI% = Asc(Left$(COPRIAT$, 1))
        ATRIP$ = Mid$(COPRIAT$, 2)
        Call SHOWCOMMAND(COPRI%, ATRIP$)
      End If
      GoTo 10
      '
    End If
    '
    UI& = UI& + 1
    XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
    If InStr(XX$, TIDOCUM$) > 4 Then
      If InStr(XX$, NUDOCU$) > 4 Then
        DOCUNU& = CVS(Left$(XX$, 4))
        GoTo 7
      End If
    End If
    '
    Screen.MousePointer = 1
    PAGINACTIVA% = 0
    For KKI% = 0 To 5
       EDITFORM.Picture1(KKI%).Visible = False
    Next KKI%
    EDITFORM.Picture1(PAGINACTIVA%).Visible = True
    EDITFORM.Picture1(PAGINACTIVA%).Refresh
    EDITFORM.VScroll1.Min = 0
    EDITFORM.VScroll1.Value = 0
    EDITFORM.VScroll1.Max = 1
    If ALTUPAGINA > EDITFORM.Frame1.Height Then
       EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
    End If
    ANCHPAGINA = EDITFORM.Picture1(PAGINACTIVA%).Width
    EDITFORM.HScroll1.Min = 0
    EDITFORM.HScroll1.Value = 0
    EDITFORM.HScroll1.Max = 1
    If ANCHPAGINA > EDITFORM.Frame1.Width Then
      EDITFORM.HScroll1.Max = (ANCHPAGINA - EDITFORM.Frame1.Width) / 100
    End If
    EDITFORM.Command3.Enabled = False
    EDITFORM.Command5.Enabled = False
    EDITFORM.Show 1
    '
    GoTo 3
    '
99  Command22.Enabled = True
    '
End Sub

Private Sub Command3_Click()
    OPXX% = ALTERNA%("Base de Datos Completa\Ordenes Pendientes")
    If Option1.Value = True Then
        Select Case OPXX%
          Case 1
            Call FINAL("*COCOMP0")
          Case 2
            Call COCOPEN        ' ORDENES PENDIENTES
            Call FINAL("*COCOPEN")
          Case Else
            Exit Sub
        End Select
        '
      ElseIf Option2.Value = True Then
        Select Case OPXX%
          Case 1
            Call FINAL("*LICOMP0")
          Case 2
            Call COCOPEN        ' ORDENES PENDIENTES
            Call FINAL("*LICOPEN")
          Case Else
            Exit Sub
        End Select
    End If
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    Call COCOPEN    ' SELECCIONAR SOLO PENDIENTES
    If Option1.Value = True Then
        Call FINAL("*CORCPRO")
      ElseIf Option2.Value = True Then
        Call FINAL("*SORCPRO")
    End If
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    Call COCOPEN    ' SELECCIONAR SOLO PENDIENTES
    If Option1.Value = True Then
        Call FINAL("*CORCMAT")
      ElseIf Option2.Value = True Then
        Call FINAL("*SORCMAT")
    End If
    Command5.Enabled = True
End Sub

Private Sub Command7_Click()
    Call OCOMNUE
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     OPENFORMS = DoEvents
   Wend
End Sub

Private Sub Form_Load()
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
End Sub

Sub COCOPEN()
   '
   Screen.MousePointer = 11
   Call COPYSTRU("OCOMDETA", "OPENDETA")
   FIN& = ULTREG&("OCOMDETA")
   JJ& = 0
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XXX$ = REGLEIDO$("OCOMDETA", U&)
     If Asc(Mid$(XXX$, 124, 1)) < 9 Then ' ordenes abiertas
       If CVS(Mid$(XXX$, 99, 4)) < CVS(Mid$(XXX$, 55, 4)) - 0.1 Then
         JJ& = JJ& + 1
         Call GRAREG("OPENDETA", XXX$, JJ&)
       End If
     End If
   Next U&
   '
   Call SETULTREG("OPENDETA", JJ&)
   Call CIERRARCH("OPENDETA")
   Screen.MousePointer = 1
   '
End Sub

Sub COCOBREC()
   '
   Dim REOCO&(8192), PUBOL$(8192), CAPEN(8192)
   '
   OPX% = COMALTER%("Elija su Opción de Control:\\Cumplimiento\Calidad\Cancelar")
   If OPX% < 1 Or OPX% > 2 Then
      Exit Sub
   End If
   '
   HAS% = HOY(HOS$)
   FEX = HAS%
   HASS$ = FECHATEX$(FEX)
   DESS$ = "01" + Mid$(HASS$, 3)
   DES% = FECHANUM(DESS$)
   VDEF$ = MKI$(DES%) + MKI$(HAS%)
   '
10 OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
   If OBX$ = "" Then
      Exit Sub
   End If
   DES% = CVI(Left$(OBX$, 2))
   HAS% = CVI(Mid$(OBX$, 3, 2))
   If HAS% < DES% Then GoTo 10
   FEX = DES%: PERIO$ = "Periodo: " + FECHATEX$(FEX)
   FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
   '
   Screen.MousePointer = 11
   If OPX% = 2 Then GoTo 50
   '
   ' INDICE DE CUMPLIMIENTO DE PROVEEDORES
   '
   J& = 0: PUCON$ = ""
   FIN& = ULTREG&("OCOMPRA")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("OCOMPRA", U&)
      FF% = CVI(Mid$(X$, 5, 2))        ' fecha de emision
      NP% = CVI(Mid$(X$, 7, 2))        ' proveedor
      STAT% = Asc(Mid$(X$, 62, 1))
      '
      If STAT% <> 9 Then               ' excluye anuladas
        If NP% > 0 Then
          If NP% <> 9999 Then            ' excluye pedidos suministro
            If FF% >= DES% Then
              If FF% <= HAS% Then
                J& = J& + 1
                If J& > 8192 Then
                  Call MENSERR(24, "Demasiados Movimientos\en el Rango de Fechas Elegido.")
                  GoTo 10
                End If
                '
                REOCO&(J&) = U&
                CAPEN(J&) = CVS(Mid$(X$, 13, 4))
                CLACO$ = Mid$(X$, 7, 2) + Left$(X$, 4) + Mid$(X$, 11, 2)
                PUCON$ = PUCON$ + CLACO$
              End If
            End If
          End If
        End If
      End If
   Next U&
   CAOCO& = J&
   '
   FIN& = ULTREG&("BOLRECEP")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("BOLRECEP", U&)
      CAREC = CVD(Mid$(X$, 89, 8))
      CLACO$ = Mid$(X$, 105, 2) + Mid$(X$, 151, 4) + Mid$(X$, 83, 2)
      UKA& = 0
      For K& = 1 To CAOCO&
        If Mid$(PUCON$, 8 * K& - 7, 8) = CLACO$ Then
          UKA& = K&
          If CAPEN(K&) > 0 Then
            PUBOL$(K&) = PUBOL$(K&) + MKS$(U&)
            CAPEN(K&) = CAPEN(K&) - CAREC
            GoTo 19
          End If
        End If
      Next K&
      '
      If UKA& > 0 Then
        PUBOL$(UKA&) = PUBOL$(UKA&) + MKS$(U&)
        CAPEN(UKA&) = CAPEN(UKA&) - CAREC
      End If
      '
19 Next U&
   '
   HOII = HOY(HOS$)
   For K& = 1 To CAOCO&
      X$ = REGLEIDO$("OCOMPRA", REOCO&(K&))
      Z0$ = REGBLAN$("COCOBREC")
      Call REPLA(Z0$, Mid$(X$, 7, 2), 1, 2)
      Call REPLA(Z0$, Left$(X$, 4), 3, 4)
      Call REPLA(Z0$, Mid$(X$, 11, 2), 7, 2)
      Call REPLA(Z0$, Mid$(X$, 13, 4), 9, 4)
      Call REPLA(Z0$, Mid$(X$, 9, 2), 13, 2)
      FENTREPRO = CVI(Mid$(X$, 9, 12))
      '
      If Len(PUBOL$(K&)) < 4 Then
           DIATRA% = 0
           If FENTREPRO < HOII Then
             DIATRA% = DIENTRE%(FENTREPRO, HOII)
           End If
           Call REPLA(Z0$, MKI$(DIATRA%), 37, 2)
           NCOCO& = NCOCO& + 1
           Call GRAREG("COCOBREC", Z0$, NCOCO&)
         Else
           For II& = 1 To Len(PUBOL$(K&)) Step 4
             Z$ = Z0$
             REBREC& = CVS(Mid$(PUBOL$(K&), II&, 4))
             Y$ = REGLEIDO$("BOLRECEP", REBREC&)
             Call REPLA(Z$, MKS$(CVD(Mid$(Y$, 89, 8))), 15, 4)
             Call REPLA(Z$, Mid$(Y$, 81, 2), 19, 2)
             Call REPLA(Z$, Mid$(Y$, 119, 16), 21, 16)
             FERECI = CVI(Mid$(Y$, 81, 2))
             DIATRA% = 0
             If FENTREPRO < FERECI Then
               DIATRA% = DIENTRE%(FENTREPRO, FERECI)
             End If
             Call REPLA(Z$, MKI$(DIATRA%), 37, 2)
             NCOCO& = NCOCO& + 1
             Call GRAREG("COCOBREC", Z$, NCOCO&)
           Next II&
      End If
      '
   Next K&
   '
   Call SETULTREG("COCOBREC", NCOCO&)
   Call CIERRARCH("COCOBREC")
   Call HEADERS("COCOBREC", "")
   Call HEADERS("COCOBREC", PERIO$)
   '
   Call FINAL("*SOCUPRO")
   Screen.MousePointer = 1
   Exit Sub
   '
' ************************************************************
   '
50 ' INDICE DE CALIDAD DE PROVEEDORES
   '
   VECLOT$ = ""
   FIN& = ULTREG&("BOLRECEP")
   JJ& = 0
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("BOLRECEP", U&)
      FF% = CVI(Mid$(X$, 81, 2))        ' fecha de emision
      NP% = CVI(Mid$(X$, 105, 2))        ' proveedor
      '
      If FF% >= DES% Then
        If FF% <= HAS% Then
          If NP% > 0 Then
            If NP% <> 9999 Then
              CANOM = CVD(Mid$(X$, 89, 8))
              If CANOM > 0 Then
                NRO& = CVS(Left$(X$, 4))
                CI% = CVI(Mid$(X$, 83, 2))
                REMI$ = Mid$(X$, 119, 16)
                CAPRO = CVD(Mid$(X$, 163, 8))
                LOTE$ = Mid$(X$, 69, 12)
                '
                Z$ = REGBLAN$("COCAREC")
                Call REPLA(Z$, MKI$(NP%), 1, 2)
                Call REPLA(Z$, MKS$(NRO&), 3, 4)
                Call REPLA(Z$, MKI$(FF%), 7, 2)
                Call REPLA(Z$, REMI$, 9, 16)
                Call REPLA(Z$, LOTE$, 25, 12)
                Call REPLA(Z$, MKI$(CI%), 37, 2)
                Call REPLA(Z$, MKS$(CANOM), 39, 4)
                Call REPLA(Z$, MKI$(1), 55, 2)
                JJ& = JJ& + 1
                Call GRAREG("COCAREC", Z$, JJ&)
                VECLOT$ = VECLOT$ + LOTE$
              End If
            End If
          End If
        End If
      End If
      '
59 Next U&
   '
   FIN& = ULTREG&("MOVISTO")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     Y$ = REGLEIDO$("MOVISTO", U&)
     If Mid$(Y$, 21, 2) = "BR" Then
       CANTI = CVD(Mid$(Y$, 88, 8)) - CVD(Mid$(Y$, 96, 8))
       If CANTI >= 0.05 Then
         LOTE$ = Mid$(Y$, 5, 12)
         If TRIM$(LOTE$) <> "" Then
           For KK& = 1 To Len(VECLOT$) Step 12
             If Mid$(VECLOT$, KK&, 12) = LOTE$ Then
               REG& = (KK& + 11) / 12
               If REG& > 0 Then
                 If REG& <= JJ& Then
                   Z$ = REGLEIDO$("COCAREC", REG&)
                   Call REPLA(Z$, MKS$(CANTI), 43, 4)
                   Call GRAREG("COCAREC", Z$, REG&)
                   GoTo 69
                 End If
               End If
             End If
           Next KK&
         End If
       End If
     End If
     '
     REFE$ = UCase$(Mid$(Y$, 43, 30))
     If Left$(REFE$, 7) = "RECHAZO" Then
       CANTI = CVD(Mid$(Y$, 96, 8)) - CVD(Mid$(Y$, 88, 8))
       If CANTI >= 0.05 Then
         LOTE$ = Mid$(Y$, 5, 12)
         If TRIM$(LOTE$) <> "" Then
           For KK& = 1 To Len(VECLOT$) Step 12
             If Mid$(VECLOT$, KK&, 12) = LOTE$ Then
               REG& = (KK& + 11) / 12
               If REG& > 0 Then
                 If REG& <= JJ& Then
                   Z$ = REGLEIDO$("COCAREC", REG&)
                   Call REPLA(Z$, MKS$(CANTI), 47, 4)
                   Call REPLA(Z$, MKI$(1), 57, 2)
                   Call GRAREG("COCAREC", Z$, REG&)
                   GoTo 69
                 End If
               End If
             End If
           Next KK&
         End If
       End If
     End If
     '
     If Left$(REFE$, 8) = "FALTANTE" Then
       CANTI = CVD(Mid$(Y$, 96, 8)) - CVD(Mid$(Y$, 88, 8))
       If CANTI >= 0.05 Then
         LOTE$ = Mid$(Y$, 5, 12)
         If TRIM$(LOTE$) <> "" Then
           For KK& = 1 To Len(VECLOT$) Step 12
             If Mid$(VECLOT$, KK&, 12) = LOTE$ Then
               REG& = (KK& + 11) / 12
               If REG& > 0 Then
                 If REG& <= JJ& Then
                   Z$ = REGLEIDO$("COCAREC", REG&)
                   Call REPLA(Z$, MKS$(CANTI), 51, 4)
                   'Call REPLA(Z$, MKI$(1), 57, 2)
                   Call GRAREG("COCAREC", Z$, REG&)
                   GoTo 69
                 End If
               End If
             End If
           Next KK&
         End If
       End If
     End If
     '
69 Next U&
   '
   Call SETULTREG("COCAREC", JJ&)
   Call CIERRARCH("COCAREC")
   Call CIERRARCH("*.*")
   Call HEADERS("COCAREC", "")
   Call HEADERS("COCAREC", PERIO$)
   '
   Call FINAL("*SOCAREC")
   Screen.MousePointer = 1
   '
End Sub

Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
End Sub

Sub OCOMNUE()                           ' NUEVA ORDEN DE COMPRA
    '
    HO% = HOY(HOS$)
    LU$ = LUDAT$
    '
    FORIPRE$ = TRIM$(CONTROL$("", "FORMOCOM"))
    If FORIPRE$ <> "" Then
        LU$ = LUDAT$
        If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
            RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
            Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Orden de Compra:\  \'" + RMCC$ + "'.")
            Exit Sub
        End If
    End If
    '
    MONUME$ = "AUTO"
    If CONTROL$("", "NUMEOCOM") = "MANUAL" Then
        MONUME$ = "MANUAL"
    End If
    '
    MONECOI% = Val(CONTROL$("OCOMPRA", "MONECOS"))
    If ECOARCH$("TAMONED", Chr$(MONECOI%)) = "" Then
       MONECOI% = 0
    End If
    '
    MODO% = COMALTER%("Puede Generar Ordenes de Compra\en Forma Directa ó sobre la Base\del Cálculo MRP\\Directa\MRP\Cancelar")
    If MODO% < 1 Or MODO% > 2 Then
        Call CIERRARCH("*.*")
        Exit Sub
    End If
    '
    AVALIDA$ = "MASTER"
    If MODO% = 2 Then AVALIDA$ = "OCOMEMI"
    '
100 Call CIERRARCH("*.*")
    NRO& = PRONUOCO&
    '
101 Call SELECHO("PROVED1")
    NPRO% = Val(VALACT1$("PROVED1"))
    If NPRO% < 1 Or ECOARCH$("PROVED1", MKI$(NPRO%)) = "" Then
        Call CIERRARCH("*.*")
        Call BAJALDISCO
        Exit Sub
    End If
    '
    X0$ = String$(10, 0) + Space$(160)
    Call REPLA(X0$, MKS$(NRO&), 1, 4)
    Call REPLA(X0$, MKI$(HO%), 5, 2)
    Call REPLA(X0$, MKI$(HO%), 7, 2)
    Call REPLA(X0$, MKI$(NPRO%), 9, 2)
    '
    DOMIPRO$ = TRIM$(CPOSCLI$((-1) * NPRO%))
    If DOMIPRO$ <> "" Then DOMIPRO$ = DOMIPRO$ + " - "
    Call REPLA(X0$, RASOCLI$((-1) * NPRO%), 11, 40)
    Call REPLA(X0$, DOMICLI$((-1) * NPRO%), 51, 40)
    Call REPLA(X0$, DOMIPRO$ + LOCACLI$((-1) * NPRO%), 91, 40)
    '
    NC1% = (-1) * NPRO%
    CPAG% = CPAGCLI%(NC1%)
    FORPAG$ = TRIM$(ECOARCH$("CONPAG", Chr$(CPAG%)))
    If FORPAG$ = "" Then FORPAG$ = "Contado Contra Entrega"
    '
11  X1$ = OBJPLA$("OCOMPRA-ENCA", X0$)
    If X1$ = "" Then GoTo 101
    '
    FECHEMI% = CVI(Mid$(X1$, 5, 2))
    If FECHEMI% > HO% Then
       Call MENSERR(24, "Fecha de Emision no Válida")
       GoTo 11
    End If
    '
    FECHENT% = CVI(Mid$(X1$, 7, 2))
    If FECHENT% < FECHEMI% Or FECHENT% < HO% Then
       Call MENSERR(24, "Fecha de Entrega no Válida")
       GoTo 11
    End If
    '
    ATEN$ = Mid$(X1$, 131, 40)
    '
    Call SETULTREG("!OCOMDETA", 0)
    '
102 TITU$ = TRIM$(Str$(NRO&)) + " - " + TRIM$(RASOCLI$(NC1%))
    VTB% = VENTABU%("CAROCO1/ARCHECO(1)=" + AVALIDA$ + "/TITUPAN=Nota de Pedido Nro. " + TITU$)
    If VTB% < 0 Then GoTo 101
    '
    If ULTREG&("!OCOMDETA") < 1 Then
       Call MENSERR(24, "Falta Ingresar Detalle Comprado")
       GoTo 102
    End If
    '
    IDENOC$ = MKS$(NRO&) + MKI$(FECHEMI%) + MKI$(NPRO%)
    For J& = 1 To ULTREG&("!OCOMDETA")
       X$ = REGLEIDO$("!OCOMDETA", J&)
       Call REPLA(X$, IDENOC$, 1, 8)
       '
       CI0% = CVI(Mid$(X$, 9, 2))
       If CODEXT$(CI0%) = "" Then
          Call MENSERR(24, "Código de Artículo no Válido\en Item Número " + TRIM$(Str$(J&)) + ".")
          GoTo 102
       End If
       UNIMAS$ = UNIMED$(CI0%)
       Call REPLA(X$, UNIMAS$, 11, 4)
       If TRIM$(Mid$(X$, 59, 4)) = "" Then
          Call REPLA(X$, UNIMED$(CI0%), 59, 4)
          Call REPLA(X$, MKS$(1), 63, 4)
       End If
       '
       CANTU = CVS(Mid$(X$, 55, 4))
       If CANTU < 0.05 Then
          Call MENSERR(24, "Falta Indicar Cantidad en Código '" + TRIM$(CODEXT$(CI0%)) + "'.")
          GoTo 102
       End If
       '
       FENTREIT% = CVI(Mid$(X$, 95, 2))
       If FENTREIT% < HO% Then
          Call REPLA(X$, MKI$(FECHENT%), 95, 2)
       End If
       '
       PREUNID# = CVD(Mid$(X$, 67, 8))
       If PREUNID# < 0.00005 Then
          UPREDES$ = ""
          Call BUSCAULTIPRE(NPRO%, CI0%, UNIMAS$, UPREDES$)
          If UPREDES$ <> "" Then
            Call REPLA(X$, UPREDES$, 59, 20)
          End If
       End If
       '
       Call GRAREG("!OCOMDETA", X$, J&)
       '
    Next J&
    '
103 VTB% = VENTABU%("CAROCO2/TITUPAN=Nota de Pedido Nro. " + TITU$)
    If VTB% < 0 Then GoTo 102
    '
    TOTANE# = 0
    For J& = 1 To ULTREG&("!OCOMDETA")
       X$ = REGLEIDO$("!OCOMDETA", J&)
       CI0% = CVI(Mid$(X$, 9, 2))
       UNIMAS$ = Mid$(X$, 11, 4)
       CANTU = CVS(Mid$(X$, 55, 4))
       PREUNID# = CVD(Mid$(X$, 67, 8))
       PORDE = CVS(Mid$(X$, 75, 4))
       FENTRE% = CVI(Mid$(X$, 95, 2))
       '
       If CODEXT$(CI0%) = "" Then
          Call MENSERR(24, "Código de Artículo no Válido\en Item Número " + TRIM$(Str$(J&)) + ".")
          GoTo 102
       End If
       If TRIM$(UNIMAS$) = "" Then
           Call REPLA(X$, UNIMED$(CI0%), 11, 4)
       End If
       If TRIM$(Mid$(X$, 59, 4)) = "" Then
          Call REPLA(X$, UNIMED$(CI0%), 59, 4)
          Call REPLA(X$, MKS$(1), 63, 4)
       End If
       '
       If CVS(Mid$(X$, 63, 4)) < 0.0001 Then
          Call REPLA(X$, MKS$(1), 63, 4)
       End If
       '
       If CANTU < 0.05 Then
          Call MENSERR(24, "Falta Indicar Cantidad en Código '" + TRIM$(CODEXT$(CI0%)) + "'.")
          GoTo 103
       End If
       '
       If PORDE < 0 Or PORDE > 99.99 Then
          Call MENSERR(24, "Porcentaje de Descuento no Válido\en Código '" + TRIM$(CODEXT$(CI0%)) + "'.")
          GoTo 103
       End If
       '
       If FENTRE% < FECHEMI% Then
          Call MENSERR(24, "Fecha de Entrega Incoherente\en Código '" + TRIM$(CODEXT$(CI0%)) + "'.")
          Call REPLA(X0$, MKI$(FECHEMI%), 9, 2)
          GoTo 103
       End If
       '
       PREUNINE# = (Int(PREUNID# * 100 * (100 - PORDE) + 0.5)) / 10000
       INETITEM# = CDbl((Int(100 * PREUNINE# * CANTU + 0.5)) / 100)
       TOTANE# = TOTANE# + INETITEM#
       '
       Call REPLA(X$, MKD$(PREUNINE#), 79, 8)
       Call REPLA(X$, MKD$(INETITEM#), 87, 8)
       Call GRAREG("!OCOMDETA", X$, J&)
       '
    Next J&
    '
    Call CIERRARCH("!OCOMDETA")
    '
    TOTIVA# = 0
    PVCL% = PIVACLI%(NC1%)
    If PVCL% <> 2 Then
      If PVCL% <> 4 Then
        If PVCL% <> 5 Then
          TOTIVA# = (Int(ALIVA * TOTANE# + 0.5)) / 100
        End If
      End If
    End If
    TOTATO# = TOTANE# + TOTIVA#
    '
' presentar registro de totales y condicion de pago
    '
123 ENCAO$ = REGBLAN$("OCOMENCA")
    TICAMBX = TICAMONE(MONECOI%)
    If TICAMBX < 0.5 Then TICAMBX = 1
    '
    Call REPLA(ENCAO$, MKS$(NRO&), 1, 4)
    FEX = FECHEMI%
    REFE$ = FECHATEX$(FEX) + " - " + RASOCLI$((-1) * NPRO%)
    Call REPLA(ENCAO$, REFE$, 5, 44)
    Call REPLA(ENCAO$, MKI$(FECHEMI%), 49, 2)
    Call REPLA(ENCAO$, MKI$(NPRO%), 51, 2)
    Call REPLA(ENCAO$, Chr$(0), 53, 1)
    Call REPLA(ENCAO$, Chr$(0), 54, 1)
    Call REPLA(ENCAO$, Chr$(MONECOI%), 83, 1)
    Call REPLA(ENCAO$, Chr$(MONECOI%), 84, 1)
    Call REPLA(ENCAO$, MKS$(TICAMBX), 85, 4)
    '
    HO% = HOY(HOS$)
    HP = 100 * Val(Left$(Time$, 2)) + Val(Mid$(Time$, 4, 2))
    Call REPLA(ENCAO$, MKI$(HO%), 55, 2)
    Call REPLA(ENCAO$, MKI$(Int(HP)), 57, 2)
    Call REPLA(ENCAO$, MKI$(USNBR%), 59, 2)
    '
    CPAG% = CPAGCLI%((-1) * NPRO%)
    Call REPLA(ENCAO$, MKI$(CPAG%), 153, 2)
    '
    Call REPLA(ENCAO$, ATEN$, 163, 30)
    '
125 PORDESCUENTO$ = TRIM$(Mid$(ENCAO, 97, 16))
    DESC$ = PORDESCUENTO$
    IDESCUEN# = TOTANE# * COEFDESCU(DESC$) / 100
    TOTANEDES# = TOTANE# - IDESCUEN#
    TOTIVA# = 0
    PVCL% = PIVACLI%(NC1%)
    If PVCL% <> 2 Then
      If PVCL% <> 4 Then
        If PVCL% <> 5 Then
          TOTIVA# = (Int(ALIVA * TOTANEDES# + 0.5)) / 100
        End If
      End If
    End If
    TOTATO# = TOTANEDES# + TOTIVA#
    Call REPLA(ENCAO$, MKD$(TOTANE#), 89, 8)
    Call REPLA(ENCAO$, MKD$(IDESCUEN#), 113, 8)
    Call REPLA(ENCAO$, MKD$(TOTIVA#), 121, 8)
    Call REPLA(ENCAO$, MKD$(TOTATO#), 145, 8)
    '
126 OBX$ = OBJPLA$("OCOMPRA-APRO", ENCAO$)
    If OBX$ = "" Then GoTo 103
    ENCAO$ = OBX$
    TOTANTE# = TOTATO#
    TOTANE# = CVD(Mid$(ENCAO, 89, 8))
    PORDESCUENTO$ = TRIM$(Mid$(ENCAO, 97, 16))
    DESC$ = PORDESCUENTO$
    IDESCUEN# = TOTANE# * COEFDESCU(DESC$) / 100
    TOTANEDES# = TOTANE# - IDESCUEN#
    TOTIVA# = 0
    PVCL% = PIVACLI%(NC1%)
    If PVCL% <> 2 Then
      If PVCL% <> 4 Then
        If PVCL% <> 5 Then
          TOTIVA# = (Int(ALIVA * TOTANEDES# + 0.5)) / 100
        End If
      End If
    End If
    TOTATO# = TOTANEDES# + TOTIVA#
    Call REPLA(ENCAO$, MKD$(TOTANE#), 89, 8)
    Call REPLA(ENCAO$, MKD$(IDESCUEN#), 113, 8)
    Call REPLA(ENCAO$, MKD$(TOTIVA#), 121, 8)
    Call REPLA(ENCAO$, MKD$(TOTATO#), 145, 8)
    If Abs(TOTANTE# - TOTATO#) > 0.005 Then
      GoTo 126
    End If
    TICAMBIO = CVS(Mid$(ENCAO$, 85, 4))
    If TICAMBIO < 1 Then
       Call REPLA(ENCAO$, MKS$(1), 85, 4)
       GoTo 126
    End If
    MONEI% = Asc(Mid$(ENCAO$, 83, 1))
    MONECOS$ = Left$(ECOARCH$("TAMONED", Mid$(ENCAO$, 83, 1)), 2)
    '
    Screen.MousePointer = 11
    Call BLOQARCH("OCOMENCA")
    Call BLOQARCH("OCOMDETA")
    Call BLOQARCH("MACONSIG")
    '
    NUOCONUE& = PRONUOCO&
    CAMBIANUM% = 0
    If NUOCONUE& <> NRO& Then
       NRO& = NUOCONUE&
       Call REPLA(ENCAO$, MKS$(NRO&), 1, 4)
       For J& = 1 To ULTREG&("!OCOMDETA")
         ZZ$ = REGLEIDO$("!OCOMDETA", J&)
         Call REPLA(ZZ$, MKS$(NRO&), 1, 4)
         Call GRAREG("!OCOMDETA", ZZ$, J&)
       Next J&
       Call CIERRARCH("!OCOMDETA")
       CAMBIANUM% = 1
    End If
    '
    RENCA& = 1 + ULTREG&("OCOMENCA")
    PRIDETA& = 1 + ULTREG&("OCOMDETA")
    ULTDETA& = PRIDETA& + ULTREG&("!OCOMDETA") - 1
    '
    ' grabar encabezado
    Call REPLA(ENCAO$, MKS$(PRIDETA&), 155, 4)
    Call REPLA(ENCAO$, MKS$(ULTDETA&), 159, 4)
    Call GRAREG("OCOMENCA", ENCAO$, RENCA&)
    Call CIERRARCH("OCOMENCA")
    '
    For J& = 1 To ULTREG&("!OCOMDETA")
       ZZ$ = REGLEIDO$("!OCOMDETA", J&)
       Call REPLA(ZZ$, MKS$(RENCA&), 125, 4)
       VAUNI$ = FORMATNUM$(CVD(Mid$(ZZ$, 79, 8)) / TICAMBIO, "F12.4")
       Call REPLA(VAUNI$, MONECOS$, 1, 2)
       Call REPLA(ZZ$, VAUNI$, 43, 12)
       Call REGAPP("OCOMDETA", ZZ$)
       '
       NOCOM$ = TRIM$(Str$(NRO&))
       While Len(NOCOM$) < 6: NOCOM$ = "0" + NOCOM$: Wend
       ITNU$ = TRIM$(Str$(J&))
       While Len(ITNU$) < 2: ITNU$ = "0" + ITNU$: Wend
       NOCOM$ = "OC-" + Left$(NOCOM$, 6) + "-" + ITNU$
    Next J&
    '
    Call CIERRARCH("OCOMDETA")
    Call BAJALDISCO
    '
' actualizar costo standard
    HOII% = HOY(HOS$)
    For J& = 1 To ULTREG&("!OCOMDETA")
      ZZ$ = REGLEIDO$("!OCOMDETA", J&)
      NUORSE& = CVS(Mid$(ZZ$, 117, 4))
      If NUORSE& < 1 Then   ' EXCLUYE ORDEN DE SERVICIO
        CI0% = CVI(Mid$(ZZ$, 9, 2))
        PREUNINE# = CVD(Mid$(ZZ$, 79, 8))
        COEFUNI = CVS(Mid$(ZZ$, 63, 4))
        If COEFUNI < 0.0001 Then COEFUNI = 1
        PREUNI1# = PREUNINE# / COEFUNI
        REMAL& = CI0% + 1
        REMAS$ = REGLEIDO$("MASTER", REMAL&)
        If CVI(Mid$(REMAS$, 123, 2)) = NPRO% Then
            PREUNI2 = PREUNI1# / TICAMBIO
            If PREUNI2 >= 0.000005 Then
              Call REPLA(REMAS$, MKS$(PREUNI2), 111, 4)
              Call REPLA(REMAS$, MKI$(FECHEMI%), 115, 2)
              Call REPLA(REMAS$, Chr$(MONEI%), 95, 1)
              Call GRAREG("MASTER", REMAS$, REMAL&)
            End If
          Else
            MONEJ% = MONECOSTO%(CI0%)
            CUNI = COSUNI(CI0%) * TICAMONE(MONEJ%)
            If CUNI >= 0.00005 Then
              If PREUNI1# > CUNI + 0.00005 Then
                Call COMUNI("ATENCION: Precio de Compra\es Superior al Costo Standard\en Código '" + TRIM$(CODEXT$(CI0%)) + "' !!!\  \Informe al Supervisor de Compras.")
              End If
            End If
        End If
      End If
    Next J&
    '
    LUENTREGA$ = TRIM$(Mid$(ENCAO$, 449, 64))
    '
    JJ& = 0
    For U& = 193 To 448 Step 64
      TTXX$ = Mid$(ENCAO$, U&, 64)
      JJ& = JJ& + 1
      Call GRAREG("!OBSERVOF", TTXX$, JJ&)
    Next U&
    Call SETULTREG("!OBSERVOF", JJ&)
    Call CIERRARCH("!OBSERVOF")
    '
    CPAG% = CVI(Mid$(ENCAO$, 153, 2))
    FORPAG$ = TRIM$(ECOARCH$("CONPAG", Chr$(CPAG%)))
    If FORPAG$ = "" Then FORPAG$ = "Contado Contra Entrega"
    '
    ' IMPRESION DE LA O/COMPRA
    '
    If CAMBIANUM% = 1 Then
      Call COMUNI("Número Ocupado.\La O/Compra fue Renumerada como Nro. " + TRIM$(Str$(NRO&)))
    End If
    '
    If TRIM$(FORIPRE$) <> "" Then
      '
      Screen.MousePointer = 11
      FECHDOC$ = FETEXTO$(FECHEMI%)
      NUMEDOC$ = TRIM$(Str$(NRO&))
      While Len(NUMEDOC$) < 6
        NUMEDOC$ = "0" + NUMEDOC$
      Wend
      TIPODOC$ = "Nota de Pedido"
      '
      CODPARAM$(1) = "COND-PAG"
      CODPARAM$(2) = "OBSERVA"
      CODPARAM$(3) = "IMP-NETO"
      CODPARAM$(4) = "POR-IVA"
      CODPARAM$(5) = "IMP-IVA"
      CODPARAM$(6) = "IMP-TOTA"
      CODPARAM$(7) = "VENDEDOR"
      CODPARAM$(8) = "LIST-PRE"
      CODPARAM$(9) = "ZONA-VEN"
      CODPARAM$(10) = "FECH-LAN"
      CODPARAM$(11) = "IMPNEDES"
      CODPARAM$(12) = "TRANSPOR"
      CODPARAM$(13) = "POR-DESC"
      CODPARAM$(14) = "IMP-DESC"
      CODPARAM$(15) = "FECH-REC"
      CODPARAM$(16) = "DOMI-TRA"
      CAPARDOC% = 16
      '
      DOCPARAM$(1) = FORPAG$
      DOCPARAM$(2) = OBSE$
      DOCPARAM$(3) = Str$(TOTANE#)
      DOCPARAM$(4) = Str$(ALIVA)
      DOCPARAM$(5) = Str$(TOTIVA#)
      DOCPARAM$(6) = Str$(TOTATO#)
      DOCPARAM$(7) = Mid$(X1$, 147, 20)
      DOCPARAM$(8) = Mid$(X1$, 67, 40)
      DOCPARAM$(9) = Mid$(X1$, 167, 40)
      DOCPARAM$(10) = Mid$(X1$, 107, 40)
      DOCPARAM$(11) = Str$(TOTANEDES#)
      DOCPARAM$(12) = TRIM$(Mid$(ENCAO$, 163, 30)) ' ATENCION SR.
      DOCPARAM$(13) = TRIM$(PORDESCUENTO$)
      DOCPARAM$(14) = Str$(IDESCUEN#)
        FEX = FECHENT%
      DOCPARAM$(15) = FECHATEX$(FEX)
      DOCPARAM$(16) = LUENTREGA$
      '
      CODTABU1$(1) = "COD-MAT"
      CODTABU1$(2) = "DES-MAT"
      CODTABU1$(3) = "UNIMED"
      CODTABU1$(4) = "CANTIDAD"
      CODTABU1$(5) = "FECH-VEN"
      CODTABU1$(6) = "REF-MAT1"
      CODTABU1$(7) = "PRE-LIST"
      CODTABU1$(8) = "PORDESCU"
      CODTABU1$(9) = "PRE-NETO"
      CODTABU1$(10) = "NET-ITEM"
      CODTABU1$(11) = "CAKILOS"
      CODTABU1$(12) = "F-PLANO"
      CODTABU1$(13) = "ORD-ITEM"
      CACOLTAB1% = 13
      '
      DESTIDOC% = NC1%
      CAITAB1% = 0
      CAITCRU% = 0
      '
      INCLUSPE% = 0
      If CONTROL$("OCOMPRA", "INCLUSPE") = "SI" Then
         INCLUSPE% = 1
      End If
      '
      REPLAORI% = 0
      If CONTROL$("OCOMPRA", "REPLAORI") = "SI" Then
        REPLAORI% = 1
      End If
      '
      For J& = 1 To ULTREG&("!OCOMDETA")
         X$ = REGLEIDO$("!OCOMDETA", J&)
         CI0% = CVI(Mid$(X$, 9, 2))
         CANTU = CVS(Mid$(X$, 55, 4))
         PREUNID# = CVD(Mid$(X$, 67, 8))
         PORDE = CVS(Mid$(X$, 75, 4))
         FEX = CVI(Mid$(X$, 95, 2))
         CONSIMAT% = 0: If TRIM$(Mid$(X$, 97, 2)) <> "" Then CONSIMAT% = 1
         '
         CODDX$ = TRIM$(CODEXT$(CI0%))
         If REPLAORI% = 1 Then
            CODDX$ = TRIM$(CODORIG$(CI0%, NC1%))
         End If
         '
         CAITAB1% = CAITAB1% + 1
         TETABU1$(1, CAITAB1%) = CODDX$
         TETABU1$(2, CAITAB1%) = DESCRIT0$(CI0%)
         TETABU1$(3, CAITAB1%) = Mid$(X$, 59, 4)
         TETABU1$(4, CAITAB1%) = TRIM$(Str$(CANTU))
         TETABU1$(5, CAITAB1%) = FECHATEX$(FEX)
         TETABU1$(6, CAITAB1%) = Mid$(X$, 97, 2)  'consigna m.prima
         TETABU1$(7, CAITAB1%) = Str$(PREUNID#)
         TETABU1$(8, CAITAB1%) = Str$(PORDE)
         TETABU1$(9, CAITAB1%) = Str$(PREUNID# * (1 - PORDE / 100))
         TETABU1$(10, CAITAB1%) = Str$(CANTU * PREUNID# * (1 - PORDE / 100))
         TETABU1$(11, CAITAB1%) = Str$(CANTU * PESUNI(CI0%))
         TETABU1$(12, CAITAB1%) = FEPLANO$(CI0%)
         TETABU1$(13, CAITAB1%) = TRIM$(Str$(J&))
         '
         If INCLUSPE% > 0 Then
           RNC$ = RECFILT$("ESPECIFI", 1, MKI$(CI0%))
           If Len(RNC$) > 0 Then
             INTERLI = Val(ATRIFORM$(FORIPRE$, "INTERLIN"))
             '
             COLUTA% = 0
             URN& = ULTREG&("ESPECIFI")
             '
             For U& = 1 To Len(RNC$) Step 4
               RENOTA& = CVS(Mid$(RNC$, U&, 4))
               If RENOTA& > 0 Then
                 If RENOTA& <= URN& Then
                   TTXY$ = RTrim$(Mid$(REGLEIDO$("ESPECIFI", RENOTA&), 3))
                   TTYY$ = REPLACAR$(TTXY$, Chr$(9), "     ")
                   If INTERLI <= 6 Then
                       CAITAB1% = CAITAB1% + 1
                       TETABU1$(2, CAITAB1%) = TTYY$
                     Else
                       COLUTA% = COLUTA% + 1
                       If COLUTA% > 2 Then COLUTA% = 1
                       If COLUTA% = 1 Then CAITAB1% = CAITAB1% + 1
                       TETABU1$(COLUTA%, CAITAB1%) = TTYY$
                   End If
                 End If
               End If
             Next U&
             '
           End If
         End If
         '
109   Next J&
      '
      Call PRINTDOC("FORMOCOM")   ' Orden de Compra
      '
    End If
    '
199 Call CIERRARCH("*.*")
    Call BAJALDISCO
    '
    GoTo 100
    '
999 Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    '
End Sub
'

Sub OCOMANU()
    '
    If COMALTER%("Realmente Desea\ANULAR Ordenes de Compra ?\\No, Abandonar\Si, Anular O/C's") <> 2 Then
      Exit Sub
    End If
    '
    HO% = HOY(HOS$)
    LU$ = LUDAT$
    '
    TITU$ = "PROCESO DE ANULACION DE ORDENES DE COMPRA"
    
2   Screen.MousePointer = 11
    OCOA& = 0: VOCO$ = ""
    FIN& = ULTREG&("OCOMDETA")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("OCOMDETA", U&)
      NOCO& = CVS(Left$(X$, 4))
      If NOCO& <> OCOA& Then
        If Asc(Mid$(X$, 124, 1)) < 9 Then ' esta abierta y pendiente
          OCOA& = NOCO&
          If CVS(Mid$(X$, 99, 4)) < CVS(Mid$(X$, 55, 4)) - 0.05 Then
            VOCO$ = VOCO$ + MKS$(NOCO&)
          End If
        End If
      End If
    Next U&
    
3   J& = 0: K& = 0: OCOA& = 0
    '
    FIN& = ULTREG&("OCOMENCA")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("OCOMENCA", U&)
      If Asc(Mid$(X$, 54, 1)) < 1 Then ' esta abierta y pendiente
        NOCO& = CVS(Left$(X$, 4))
        For KKL& = 1 To Len(VOCO$) Step 4
          If CVS(Mid$(VOCO$, KKL&, 4)) = NOCO& Then GoTo 5
        Next KKL&
        GoTo 9
        '
5       Y$ = Left$(X$, 48) + MKS$(U&)
        J& = J& + 1
        Call GRAREG("INDIOCOM", Y$, J&)
      End If
9   Next U&
    Call SETULTREG("INDIOCOM", J&)
    Call CIERRARCH("INDIOCOM")
    Call BAJALDISCO
    Screen.MousePointer = 1
    '
    If J& < 1 Then
      Call MENSERR(24, "No hay Ordenes de Compra\Pendientes que Puedan Anularse.")
      Screen.MousePointer = 1
      GoTo 999
    End If
    '
    INDI% = 1
    Screen.MousePointer = 1
    '
10  YY$ = REGSEL$("INDIOCOM/Indice de Ordenes de Compra Pendientes para Anulación")
    If Len(YY$) <= 4 Then GoTo 999
    '
    NUOCO& = CVS(Left$(YY$, 4))
    RENCA& = CVS(Mid$(YY$, 49, 4))
    If RENCA& < 1 Or RENCA& > ULTREG&("OCOMENCA") Then
      Call MENSERR(24, "Error de Indices - Consulte")
      GoTo 999
    End If
    XX$ = REGLEIDO$("OCOMENCA", RENCA&)
    If CVS(Left$(XX$, 4)) <> NUOCO& Then
      Call MENSERR(24, "Error de Indices - Consulte")
      GoTo 999
    End If
    '
    Screen.MousePointer = 11
    FIN& = ULTREG&("OCOMENCA")
    For J& = 1 To FIN&
      Call TRACE(20, J&, FIN&)
      XX$ = REGLEIDO$("OCOMENCA", J&)
      If CVS(Left$(XX$, 4)) = NUOCO& Then
        Call REPLA(XX$, Chr$(9), 54, 1)
        Call GRAREG("OCOMENCA", XX$, J&)
      End If
    Next J&
    Call CIERRARCH("OCOMENCA")
    '
    FIN& = ULTREG&("OCOMDETA")
    For J& = 1 To FIN&
      Call TRACE(20, J&, FIN&)
      XX$ = REGLEIDO$("OCOMDETA", J&)
      If CVS(Left$(XX$, 4)) = NUOCO& Then
        If CVS(Mid$(XX$, 99, 4)) < CVS(Mid$(XX$, 55, 4)) - 0.05 Then
          Call REPLA(XX$, Chr$(9), 124, 1)
          Call GRAREG("OCOMDETA", XX$, J&)
        End If
      End If
    Next J&
    Call CIERRARCH("OCOMDETA")
    '
    J& = 0
    For I& = 1 To ULTREG&("INDIOCOM")
      XX$ = REGLEIDO$("INDIOCOM", I&)
      If CVS(Left$(XX$, 4)) <> NUOCO& Then
        J& = J& + 1
        Call GRAREG("INDIOCOM", XX$, J&)
      End If
    Next I&
    Call SETULTREG("INDIOCOM", J&)
    '
    ' CANCELAR RESERVAS
    '
    JJ& = 0
    For J& = 1 To ULTREG&("MACONSIG")
      XX$ = REGLEIDO$("MACONSIG", J&)
      If Val(Mid$(XX$, 24, 6)) <> NUOCO& Then
           JJ& = JJ& + 1
           Call GRAREG("MACONSIG", XX$, JJ&)
         Else
           CAENTRE# = CVD(Mid$(XX$, 63, 8))
           CADEVUE# = CVD(Mid$(XX$, 111, 8))
           If Abs(CAENTRE# - CADEVUE#) >= 0.05 Then
             JJ& = JJ& + 1
             Call GRAREG("MACONSIG", XX$, JJ&)
           End If
      End If
    Next J&
    Call SETULTREG("MACONSIG", JJ&)
    '
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Call FRESHECHO("INDIOCOM")
    Screen.MousePointer = 1
    GoTo 10
    '
999 Call CIERRARCH("*.*")
    '
End Sub
'
Sub BUSCAULTIPRE(NPRO%, CIIX%, UNIMAS$, UPREDES$)
   '
   RFF$ = RECFILT$("OCOMDETA", 3, MKI$(NPRO%))
   If Len(RFF$) >= 4 Then
     For U& = 1 To Len(RFF$) Step 4
       REGO& = CVS(Mid$(RFF$, U&, 4))
       YX$ = REGLEIDO$("OCOMDETA", REGO&)
       If CVI(Mid$(YX$, 7, 2)) = NPRO% Then
         If CVI(Mid$(YX$, 9, 2)) = CIIX% Then
           If TRIM$(Mid$(YX$, 11, 4)) = TRIM$(UNIMAS$) Then
             UPREDES$ = Mid$(YX$, 59, 20)
           End If
         End If
       End If
     Next U&
   End If
   '
End Sub

Function PRONUOCO&()
   '
   PRONU& = 1
   REGXX& = ULTREG&("OCOMENCA")
   If REGXX& > 0 Then
     PRONU& = 1 + CVS(Left$(REGLEIDO$("OCOMENCA", REGXX&), 4))
   End If
   If PRONU& <= REGXX& Then
     Screen.MousePointer = 11
     For U& = 1 To REGXX&
       X0$ = REGLEIDO$("OCOMENCA", U&)
       If CVS(Left$(X0$, 4)) >= PRONU& Then
         PRONU& = 1 + CVS(Left$(X0$, 4))
       End If
     Next U&
     Screen.MousePointer = 1
   End If
   '
   PRONUOCO& = PRONU&
   '
End Function

Sub SELCOMRAFE()
   '
   HAS% = HOY(HOS$)
   FEX = HAS%
   HASS$ = FECHATEX$(FEX)
   DESS$ = "01" + Mid$(HASS$, 3)
   DES% = FECHANUM(DESS$)
   VDEF$ = MKI$(DES%) + MKI$(HAS%)
   '
10 OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
   If OBX$ = "" Then
      Exit Sub
   End If
   DES% = CVI(Left$(OBX$, 2))
   HAS% = CVI(Mid$(OBX$, 3, 2))
   If HAS% < DES% Then GoTo 10
   FEX = DES%: PERIO$ = "Periodo: " + FECHATEX$(FEX)
   FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
   '
   Screen.MousePointer = 11
   FIN& = ULTREG&("OCOMDETA")
   JJ& = 0
   REBLA$ = REGBLAN$("COPRECOM")
   Call HEADERS("COPRECOM", "")
   Call HEADERS("COPRECOM", "Periodo: " + PERIO$)
   '
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XXX$ = REGLEIDO$("OCOMDETA", U&)
     FECHEMI% = CVI(Mid$(XXX$, 5, 2))
     '
     STAT% = Asc(Mid$(XXX$, 124, 1))
     If STAT% <> 9 Then           ' no esta anulada
       If FECHEMI% >= DES% Then
         If FECHEMI% <= HAS% Then
           '
           NPRO% = CVI(Mid$(XXX$, 7, 2))
           CIXI% = CVI(Mid$(XXX$, 9, 2))
           DOPE$ = ""
           NOPE% = CVI(Mid$(XXX$, 103, 2))
           If NOPE% > 0 Then
             DOPE$ = DESCRIOP$(CIXI%, NOPE%)
           End If
           NUOC& = CVS(Left$(XXX$, 4))
           UMEDI$ = Mid$(XXX$, 59, 4)
           If TRIM$(UMEDI$) = "" Then
             UMEDI$ = UNIMED$(CIXI%)
           End If
           CACOM = CVS(Mid$(XXX$, 55, 4))
           PREUNI# = CVD(Mid$(XXX$, 79, 8))
           '
           YYY$ = REBLA$
           Call REPLA(YYY$, MKI$(CIXI%), 1, 2)
           Call REPLA(YYY$, MKI$(FECHEMI%), 3, 2)
           Call REPLA(YYY$, MKI$(NPRO%), 5, 2)
           Call REPLA(YYY$, MKI$(CIXI%), 7, 2)
           Call REPLA(YYY$, MKI$(FECHEMI%), 9, 2)
           Call REPLA(YYY$, DOPE$, 11, 22)
           Call REPLA(YYY$, MKS$(NUOC&), 33, 4)
           Call REPLA(YYY$, UMEDI$, 37, 4)
           Call REPLA(YYY$, MKS$(CACOM), 41, 4)
           Call REPLA(YYY$, MKD$(PREUNI#), 49, 8)
           ITOCOM# = PREUNI# * CACOM
           Call REPLA(YYY$, MKD$(ITOCOM#), 57, 8)
           JJ& = JJ& + 1
           Call GRAREG("COPRECOM", YYY$, JJ&)
         End If
       End If
     End If
     '
   Next U&
   '
   Call SETULTREG("COPRECOM", JJ&)
   Call CIERRARCH("COPRECOM")
   Screen.MousePointer = 1
   '
End Sub
'
Sub DESANULOC()
   '
3  Call COPYSTRU("INDIOCOM", "IOCANUL")
   J& = 0: K& = 0: OCOA& = 0
   '
   FIN& = ULTREG&("OCOMENCA")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("OCOMENCA", U&)
      If Asc(Mid$(X$, 54, 1)) = 9 Then ' esta anulada
5       Y$ = Left$(X$, 48) + MKS$(U&)
        J& = J& + 1
        Call GRAREG("IOCANUL", Y$, J&)
      End If
9   Next U&
    Call SETULTREG("IOCANUL", J&)
    Call CIERRARCH("IOCANUL")
    Call BAJALDISCO
    Screen.MousePointer = 1
    '
    If J& < 1 Then
      Call MENSERR(24, "No hay Ordenes de Compra\Anuladas que Puedan Restaurarse.")
      Screen.MousePointer = 1
      GoTo 999
    End If
    '
    INDI% = 1
    Screen.MousePointer = 1
    '
10  YY$ = REGSEL$("IOCANUL/Indice de Ordenes de Compra Anuladas para Restauración")
    If Len(YY$) <= 4 Then GoTo 999
    '
    NUOCO& = CVS(Left$(YY$, 4))
    RENCA& = CVS(Mid$(YY$, 49, 4))
    If RENCA& < 1 Or RENCA& > ULTREG&("OCOMENCA") Then
      Call MENSERR(24, "Error de Indices - Consulte")
      GoTo 999
    End If
    XX$ = REGLEIDO$("OCOMENCA", RENCA&)
    If CVS(Left$(XX$, 4)) <> NUOCO& Then
      Call MENSERR(24, "Error de Indices - Consulte")
      GoTo 999
    End If
    '
    Screen.MousePointer = 11
    FIN& = ULTREG&("OCOMENCA")
    For J& = 1 To FIN&
      Call TRACE(20, J&, FIN&)
      XX$ = REGLEIDO$("OCOMENCA", J&)
      If CVS(Left$(XX$, 4)) = NUOCO& Then
        Call REPLA(XX$, Chr$(0), 54, 1)
        Call GRAREG("OCOMENCA", XX$, J&)
      End If
    Next J&
    Call CIERRARCH("OCOMENCA")
    '
    FIN& = ULTREG&("OCOMDETA")
    For J& = 1 To FIN&
      Call TRACE(20, J&, FIN&)
      XX$ = REGLEIDO$("OCOMDETA", J&)
      If CVS(Left$(XX$, 4)) = NUOCO& Then
        Call REPLA(XX$, Chr$(0), 124, 1)
        Call GRAREG("OCOMDETA", XX$, J&)
      End If
    Next J&
    Call CIERRARCH("OCOMDETA")
    '
    J& = 0
    For I& = 1 To ULTREG&("IOCANUL")
      XX$ = REGLEIDO$("IOCANUL", I&)
      If CVS(Left$(XX$, 4)) <> NUOCO& Then
        J& = J& + 1
        Call GRAREG("IOCANUL", XX$, J&)
      End If
    Next I&
    Call SETULTREG("IOCANUL", J&)
    '
    GoTo 10
    '
999 Call CIERRARCH("OCOMENCA")
    Call CIERRARCH("OCOMDETA")
    Call BAJALDISCO
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Hide
End Sub
