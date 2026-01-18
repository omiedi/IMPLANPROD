VERSION 5.00
Begin VB.Form OPRETRAS 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Remitos de Traslado - Opciones"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5160
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
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
      Height          =   855
      Left            =   4095
      Picture         =   "OPRETRAS.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Terminar - Salir de la Aplicación"
      Top             =   315
      Width           =   855
   End
   Begin VB.CommandButton Command7 
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
      Height          =   855
      Left            =   4095
      Picture         =   "OPRETRAS.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
      Top             =   1260
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Setup - Configuración"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Index           =   1
      Left            =   210
      TabIndex        =   4
      Top             =   1680
      Width           =   3690
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   525
         ScaleHeight     =   585
         ScaleWidth      =   2580
         TabIndex        =   5
         Top             =   420
         Width           =   2640
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
            Height          =   600
            Left            =   0
            Picture         =   "OPRETRAS.frx":0454
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Acceso Directo a Configuración de Funcionaliento"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Configuracion de Funcionamiento"
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
            TabIndex        =   7
            ToolTipText     =   "Click Aqui Importa Archivo GETRAMO"
            Top             =   105
            Width           =   1635
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Operaciones Especiales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Index           =   0
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   3690
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   525
         ScaleHeight     =   585
         ScaleWidth      =   2595
         TabIndex        =   1
         Top             =   420
         Width           =   2655
         Begin VB.CommandButton Command25 
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
            Picture         =   "OPRETRAS.frx":0896
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Anulación General de Remito / Orden de Despacho Emitidos"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Anulación de Remito de Traslado"
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
            TabIndex        =   3
            ToolTipText     =   "Click Aqui Importa Archivo GETRAMO"
            Top             =   105
            Width           =   1740
         End
      End
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   3560
      Picture         =   "OPRETRAS.frx":0BA0
      ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
      Top             =   2430
      Width           =   2010
   End
End
Attribute VB_Name = "OPRETRAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    '
    command1.Enabled = False
    OPREMITRA.Show 1
    command1.Enabled = True
    '
End Sub


Private Sub Command2_Click()
    '
    Call CIERRARCH("*.*")
    Hide
    '
End Sub

Private Sub Command25_Click()
    Command25.Enabled = False
    Dim CISL%(2048), CANSL#(2048), DEPSL%(2048)
    '
    CMO$ = TRIM$(CONTROL$("", "CODRETRA"))
    If Len(CMO$) <> 2 Then
       CMO$ = "RT"
    End If
    '
    Call CARSELMESA
    '
10  Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) = "" Then
       GoTo 99
    End If
    CORRESQ$ = MESANO$(MESASEL$, DES%, HAS%)
    '
    Call FRESHECHO("SELREMI")
    Screen.MousePointer = 11
    FIN& = ULTREG&("MOVISTO")
    J& = 0: REMIA$ = ""
    '
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = REGLEIDO$("MOVISTO", U&)
       FF% = CVI(Left$(X$, 2))
       If FF% >= DES% Then
          If FF% <= HAS% Then
             REMI$ = UCase$(Mid$(X$, 23, 10))
             If Val(Mid$(REMI$, 4)) > 0 Then
               If REMI$ <> REMIA$ Then
                 REMIA$ = REMI$
                 TIMO1$ = Mid$(X$, 21, 2)
                 If TIMO1$ = CMO$ Then
                   CLI% = CVI(Mid$(X$, 84, 2))
                   Z$ = Space$(40) + String$(8, 0)
                   Call REPLA(Z$, REMI$, 1, 10)
                   Call REPLA(Z$, RASOCLI$(CLI%), 11, 30)
                   Call REPLA(Z$, MKS$(U&), 41, 4)
                   Call REPLA(Z$, MKI$(FF%), 45, 2)
                   J& = J& + 1
                   Call GRAREG("SELREMI", Z$, J&)
                 End If
               End If
             End If
          End If
       End If
    Next U&
    '
    Call SETULTREG("SELREMI", J&)
    Call CIERRARCH("SELREMI")
    Screen.MousePointer = 1
    '
    If J& < 1 Then
       Call MENSERR(24, "No se Encontraron Remitos\en el Período Elegido.")
       GoTo 10
    End If
    '
    Call SELECHO("SELREMI")
    REMISEL$ = VALACT1$("SELREMI")
    If TRIM$(REMISEL$) = "" Then
       GoTo 10
    End If
    '
    RERE$ = REGACT$("SELREMI")
    PRIREG& = CVS(Mid$(RERE$, 41, 4))
    CLIE$ = Mid$(RERE$, 11, 30)
    FEX = CVI(Mid$(RERE$, 45, 2))
    '
    If PRIREG& < 1 Then GoTo 10
    '
    TXT$ = "Remito Número " + TRIM$(REMISEL$) + "\Cliente: " + TRIM$(CLIE$) + "\Fecha: " + FECHATEX$(FEX) + "\   \\Abandonar\Anular"
    ALTE% = COMALTER%(TXT$)
    If ALTE% <> 2 Then GoTo 10
    '
    Call BLOQARCH("MOVISTO")
    '
    Screen.MousePointer = 11
    FIN& = ULTREG&("MOVISTO")
    KKJ% = 0: FECHA% = 0
    For U& = PRIREG& To FIN&
    Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("MOVISTO", U&)
      If TRIM$(Mid$(X$, 23, 10)) = REMISEL$ Then
        If FECHA% = 0 Then
           FECHA% = CVI(Left$(X$, 2))
           CLI% = CVI(Mid$(X$, 84, 2))
        End If
        CII% = CVI(Mid$(X$, 3, 2))
        DEPO% = Asc(Mid$(X$, 83, 1))
        CANTU# = CVD(Mid$(X$, 88, 8)) - CVD(Mid$(X$, 96, 8))
        For KKI% = 1 To KKJ%
          If CISL%(KKI%) = CII% Then
            If DEPSL%(KKI%) = DEPO% Then
               CANSL#(KKI%) = CANSL#(KKI%) + CANTU#
               GoTo 20
            End If
          End If
        Next KKI%
        KKJ% = KKJ% + 1
        CISL%(KKJ%) = CII%
        DEPSL%(KKJ%) = DEPO%
        CANSL#(KKJ%) = CANTU#
      End If
20  Next U&
    '
    For KKI% = 1 To KKJ%
      If Abs(CANSL#(KKI%)) >= 0.05 Then
        CI% = CISL%(KKI%)
        DEPO% = DEPSL%(KKI%)
        LOTE$ = ""
        CMOV$ = CMO$
        DOPRI$ = REMISEL$
        DOSEC$ = ""
        REFE$ = "Anulac.Remito " + REMISEL$
        VUNI# = 0
        MONE$ = ""
        NC% = CLI%
        CANTI = (-1) * CANSL#(KKI%)
        Call MOVISTO(FECHA%, CI%, LOTE$, CMOV$, DOPRI$, DOSEC$, REFE$, VUNI#, MONE$, NC%, DEPO%, CANTI)
      End If
    Next KKI%
    '
99  Call CIERRARCH("MOVISTO")
    Call CIERRARCH("*.*")
    Call LIBARCH("*.*")
    Reset
    Screen.MousePointer = 1
    Command25.Enabled = True
    '
End Sub

Private Sub Command7_Click()
    Call HELPONLINE("REMITRAS")
End Sub
