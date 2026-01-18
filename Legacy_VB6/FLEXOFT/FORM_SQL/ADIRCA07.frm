VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form ADIRCA07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0D0D0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "C.R.P. - Programación de Carga de Máquinas"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5190
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   5190
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   1
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   1215
      TabIndex        =   25
      Top             =   0
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   3060
      Left            =   4320
      ScaleHeight     =   3030
      ScaleWidth      =   1350
      TabIndex        =   21
      Top             =   0
      Width           =   1380
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
         Height          =   640
         Left            =   105
         Picture         =   "ADIRCA07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   250
         Width           =   650
      End
      Begin VB.CommandButton Command29 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "ADIRCA07.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   890
         Width           =   650
      End
      Begin VB.CommandButton Command34 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   105
         Picture         =   "ADIRCA07.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   1530
         Width           =   650
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0D0D0&
      Caption         =   "APLICACIONES COMPLEMENTARIAS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2640
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4110
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   1365
         ScaleHeight     =   615
         ScaleWidth      =   2295
         TabIndex        =   8
         Top             =   1785
         Width           =   2325
         Begin VB.CommandButton Command35 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   630
            Left            =   0
            Picture         =   "ADIRCA07.frx":075E
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Acceso Directo a Calculo de Necesidades de Materiales"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Programación de Materiales MRP-II"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   630
            TabIndex        =   10
            Top             =   105
            Width           =   1800
         End
      End
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   1050
         ScaleHeight     =   615
         ScaleWidth      =   2295
         TabIndex        =   17
         Top             =   1365
         Width           =   2325
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
            Height          =   630
            Left            =   0
            Picture         =   "ADIRCA07.frx":0A68
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Acceso Directo a Gestión de Compras"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Gestión de Compras"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   720
            TabIndex        =   18
            Top             =   120
            Width           =   1485
         End
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   735
         ScaleHeight     =   615
         ScaleWidth      =   2295
         TabIndex        =   11
         Top             =   945
         Width           =   2325
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
            Height          =   630
            Left            =   0
            Picture         =   "ADIRCA07.frx":0D72
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Acceso Directo a Informes de Producción"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Informes de Producción"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   840
            TabIndex        =   13
            Top             =   0
            Width           =   1320
         End
      End
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   420
         ScaleHeight     =   615
         ScaleWidth      =   2295
         TabIndex        =   14
         Top             =   525
         Width           =   2325
         Begin VB.CommandButton Command24 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   630
            Left            =   0
            Picture         =   "ADIRCA07.frx":107C
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Acceso Directo a Ordenes de Fabricación"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Ordenes de Fabricación"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   840
            TabIndex        =   16
            Top             =   0
            Width           =   1695
         End
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Listados de Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2640
         Left            =   2625
         TabIndex        =   1
         Top             =   4830
         Width           =   4005
         Begin VB.CommandButton Command23 
            Caption         =   "Inventarios"
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
            Left            =   210
            TabIndex        =   7
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command22 
            Caption         =   "Negativos"
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
            Left            =   2205
            TabIndex        =   6
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command21 
            Caption         =   "Consignación"
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
            Left            =   210
            TabIndex        =   5
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command20 
            Caption         =   "Críticos"
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
            Left            =   2205
            TabIndex        =   4
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command17 
            Caption         =   "Estadísticas"
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
            Left            =   210
            TabIndex        =   3
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command16 
            Caption         =   "Análisis A-B-C"
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
            Left            =   2205
            TabIndex        =   2
            Top             =   1890
            Width           =   1590
         End
      End
   End
   Begin VB.ListBox SORTLIST 
      Height          =   255
      Left            =   3885
      Sorted          =   -1  'True
      TabIndex        =   20
      Top             =   1890
      Visible         =   0   'False
      Width           =   1695
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   4020
      OleObjectBlob   =   "ADIRCA07.frx":1386
      Top             =   2100
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   4080
      OleObjectBlob   =   "ADIRCA07.frx":15BA
      TabIndex        =   26
      Top             =   420
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "ADIRCA07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command14_Click()
    Command14.Enabled = False
    Call CONECRUN("ORCOMP07", "Gestion de Ordenes de Compra")
    Command14.Enabled = True
End Sub

Private Sub Command24_Click()
    Command24.Enabled = False
    Call CONECRUN("ORFABR07", "Ordenes de Fabricación")
    Command24.Enabled = True
End Sub

Private Sub Command25_Click()
    Command25.Enabled = False
    Call CONECRUN("IFABRI07", "Informes de Producción")
    Command25.Enabled = True
End Sub

Private Sub Command29_Click()
    Call HELPONLINE("CARGMAQ")
End Sub

Private Sub Command34_Click()
    Call ACCDIR("CARGMAQ")
End Sub

Private Sub Command35_Click()
    Command35.Enabled = False
    Call CONECRUN("PROMRP07", "Programacion de Materiales - MRP-II")
    Command35.Enabled = True
End Sub

Private Sub Form_Load()
   Call APLICACIONSKINS(Me, Picture1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call BAJALDISCO
End Sub

Sub INFOPCUMPLI()      ' OPERACIONES CUMPLIDAS
     '
100  NROOF$ = ""

1405 Call FRESHECHO("ECORDEP")
     If ULTREG&("ECORDEP") < 1 Then
         Call MENSERR(24, "No hay Ordenes de Fabricación Pendientes.")
         Exit Sub
     End If
     '
     Call SELECHO("ECORDEP/Ordenes de Trabajo en Curso")
1410 NROOF$ = AJUSTI$(VALACT1$("ECORDEP"), 12)
     If TRIM$(NROOF$) = "" Then
         Exit Sub
     End If
     '
     Screen.MousePointer = 11
     BVOF% = BINVAL%(AJUSTI$(NROOF$, 12))
     RGFSX$ = RECFILT$("OPERFAB", 6, MKI$(BVOF%))
     RGFSY$ = ""
     For KJ& = 1 To Len(RGFSX$) Step 4
       REOPE& = CVS(Mid$(RGFSX$, KJ&, 4))
       If REOPE& > 0 Then
         If REOPE& <= ULTREG&("OPERFAB") Then
           XX$ = REGLEIDO$("OPERFAB", REOPE&)
           If Left$(XX$, 12) = NROOF$ Then
             If Mid$(XX$, 177, 6) = String$(6, 0) Then ' aun pendiente
               RGFSY$ = RGFSY$ + MKS$(REOPE&)
             End If
           End If
         End If
       End If
     Next KJ&
     '
     If Len(RGFSY$) < 4 Then
       For KJ& = 1 To ULTREG&("OPERFAB")
         XX$ = REGLEIDO$("OPERFAB", KJ&)
         If Left$(XX$, 12) = NROOF$ Then
           If Mid$(XX$, 177, 6) = String$(6, 0) Then ' aun pendiente
             Call REPLA(XX$, MKI$(BVOF%), 65, 2)
             Call GRAREG("OPERFAB", XX$, KJ&)
             RGFSY$ = RGFSY$ + MKS$(KJ&)
           End If
         End If
       Next KJ&
       Call CIERRARCH("OPERFAB")
     End If
     '
     If Len(RGFSY$) < 4 Then
         Call MENSERR(24, "No hay Operaciones Programadas\para esa Orden de Fabricación.")
         Screen.MousePointer = 1
         GoTo 1405
     End If
     '
     j% = 0
     Call FRESHECHO("!OPERINF")
     SORTLIST.Clear
     For K% = 1 To Len(RGFSY$) Step 4
        REGI& = CVS(Mid$(RGFSY$, K%, 4))
        XX$ = REGLEIDO$("OPERFAB", REGI&)
        NOPE% = CVI(Mid$(XX$, 15, 2))
        TSOR$ = CSTRING$(Mid$(XX$, 15, 2), 1, 6, 8, 2) + Str$(REGI&)
        SORTLIST.AddItem TSOR$
     Next K%
     '
     For K% = 1 To SORTLIST.ListCount
        REGI& = Val(Mid$(SORTLIST.List(K% - 1), 7))
        XX$ = REGLEIDO$("OPERFAB", REGI&)
        CAPROD# = CVD(Mid$(XX$, 79, 8))
        CAREA# = CVD(Mid$(XX$, 87, 8))
        SALDO# = CAPROD# - CAREA#: If SALDO# < 0 Then SALDO# = 0
        '
        If SALDO# > 0.05 Then
            '
            CI% = CVI(Mid$(XX$, 13, 2))
            NOPE% = CVI(Mid$(XX$, 15, 2))
            DOPE$ = Mid$(XX$, 17, 48)
            CAINF# = SALDO#
            MAQUI$ = Mid$(XX$, 73, 6)
            FEOP% = CVI(Mid$(XX$, 139, 2))
            HOOP% = CVI(Mid$(XX$, 141, 2))
            MALIB$ = "S"
            TITOT = CVS(Mid$(XX$, 131, 4))
            '
            j% = j% + 1
            YY$ = MKI$(j%) + Space$(58) + MKS$(REGI&)
            NOPX$ = AJUSTD$(Str$(NOPE%), 5) + " - " + AJUSTI$(TRIM$(Left$(DOPE$, 31)) + " " + MAQUI$, 38)
            NOPX$ = NOPX$ + NUMSTRI$(SALDO#, 12, 1, 0)
            Call REPLA(YY$, NOPX$, 3, 58)
            JJ& = j%
            Call GRAREG("!OPERINF", YY$, JJ&)
            '
        End If
     Next K%
     '
     If j% < 1 Then
         Call MENSERR(24, "No hay Operaciones Programadas\para esa Orden de Fabricación.")
         Screen.MousePointer = 1
         GoTo 1405
     End If
     '
     Call SETULTREG("!OPERINF", JJ&)
     Call CIERRARCH("!OPERINF")
     '
     Screen.MousePointer = 1
     '
1415 Call SELECHO("!OPERINF/Operaciones Programadas para O/F " + NROOF$)
     NORDE& = Val(VALACT1$("!OPERINF"))
     If NORDE& < 1 Then
         GoTo 1405
     End If
     '
     XX$ = REGLEIDO$("!OPERINF", NORDE&)
     REGI& = CVS(Mid$(XX$, 61, 4))
     Call CARGAINF(REGI&)
     GoTo 1405
     '
End Sub
'
Sub CARGAINF(REGI&)
     '
     HOII% = HOY(HOS$)
     '
     ZZ$ = REGLEIDO$("OPERFAB", REGI&)
        CAPROD# = CVD(Mid$(ZZ$, 79, 8))
        CAREA# = CVD(Mid$(ZZ$, 87, 8))
        SALDO# = CAPROD# - CAREA#: If SALDO# < 0 Then SALDO# = 0
     CAINF# = SALDO#
     '
10   PROX$ = "Unidades Cumplidas"
     TITX$ = TRIM$(Left$(ZZ$, 12))
     TITX$ = TITX$ + " - Operación: " + TRIM$(Mid$(ZZ$, 17, 48))
     CAIN1# = Val(InputBox$(PROX$, TITX$, FORMATNUM$(CAINF#, "F10.1")))
     If CAIN1# < 0.05 Then Exit Sub
     If CAIN1# > CAINF# + 0.06 Then GoTo 10
     If CAIN1# > CAINF# Then CAIN1# = CAINF#
     '
     'TTXX$ = TRIM$(Left$(ZZ$, 12))
     'TTXX$ = TTXX$ + " - Operación: " + TRIM$(Mid$(ZZ$, 17, 48))
     'TTXX$ = TTXX$ + " - Cant: " + TRIM$(FORMATNUM$(CAINF#, "F7.0"))
     'OPX% = COMALTER%(TTXX$ + "\\Dejar Pendiente\Marcar Cumplida")
     'If OPX% <> 2 Then Exit Sub
     '
     HOOP% = Int(HORANUM(Time$))
     '
     FEBA% = HOII%
     HOBA% = Int(HORANUM(Time$))
     USUA% = USNBR% Mod 100
     DBAJA$ = MKI$(FEBA%) + MKI$(HOBA%) + Chr$(USUA%) + Chr$(1) + "Parte Fabric."
     '
     PRANT# = CVD(Mid$(ZZ$, 87, 8))
     Call REPLA(ZZ$, MKD$(PRANT# + CAIN1#), 87, 8)
     SALNUE# = CVD(Mid$(ZZ$, 79, 8)) - CVD(Mid$(ZZ$, 87, 8)): If SALNUE# < 0 Then SALNUE# = 0
     Call REPLA(ZZ$, MKD$(SALNUE#), 95, 8)
     Call REPLA(ZZ$, MKI$(HOII%), 157, 2)
     Call REPLA(ZZ$, MKI$(HOOP%), 159, 2)
     Call REPLA(ZZ$, MKS$(TEJEC), 161, 4)
     Call REPLA(ZZ$, LEGOPER$, 165, 8)
     If SALNUE# < 0.05 Then
       Call REPLA(ZZ$, DBAJA$, 177, 22)
     End If
     Call GRAREG("OPERFAB", ZZ$, REGI&)
     Call CIERRARCH("OPERFAB")
     '
1499 Call CIERRARCH("*.*")
     Call BAJALDISCO
     ATENTI% = 1 ' PARA RECALCULAR CARGA DE MáQUINAS
     '
End Sub
'

