VERSION 5.00
Begin VB.Form FESTSU07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0EEC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control Estadístico de Consumos"
   ClientHeight    =   4830
   ClientLeft      =   390
   ClientTop       =   300
   ClientWidth     =   9210
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   9210
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0EEC0&
      Caption         =   "Ultimo Cálculo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4450
      Left            =   210
      TabIndex        =   23
      Top             =   210
      Width           =   4845
      Begin VB.Frame Frame9 
         BackColor       =   &H00C0EEC0&
         Caption         =   "Período"
         Height          =   1065
         Left            =   210
         TabIndex        =   27
         Top             =   420
         Width           =   3165
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Desde:"
            Height          =   225
            Left            =   210
            TabIndex        =   31
            Top             =   380
            Width           =   750
         End
         Begin VB.Label label1 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1050
            TabIndex        =   30
            Top             =   315
            Width           =   1935
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Hasta:"
            Height          =   225
            Left            =   210
            TabIndex        =   29
            Top             =   690
            Width           =   750
         End
         Begin VB.Label Label9 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1050
            TabIndex        =   28
            Top             =   630
            Width           =   1935
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0EEC0&
         Caption         =   "Meses"
         Height          =   1065
         Left            =   3465
         TabIndex        =   25
         Top             =   420
         Width           =   1170
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   240
            TabIndex        =   26
            Top             =   315
            Width           =   675
         End
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1320
         Left            =   1260
         Sorted          =   -1  'True
         TabIndex        =   24
         Top             =   2940
         Width           =   3165
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1260
         TabIndex        =   36
         Top             =   1725
         Width           =   3130
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "F.Cálculo:"
         Height          =   225
         Index           =   0
         Left            =   105
         TabIndex        =   35
         Top             =   1780
         Width           =   1065
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Sucursales:"
         Height          =   225
         Index           =   1
         Left            =   105
         TabIndex        =   34
         Top             =   2265
         Width           =   1065
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1260
         TabIndex        =   33
         Top             =   2240
         Width           =   3135
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Clientes:"
         Height          =   225
         Index           =   2
         Left            =   105
         TabIndex        =   32
         Top             =   3045
         Width           =   1065
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00400040&
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   5160
      ScaleHeight     =   1035
      ScaleWidth      =   4185
      TabIndex        =   18
      Top             =   360
      Width           =   4245
      Begin VB.CommandButton command1 
         Caption         =   "Exit"
         Height          =   700
         Left            =   3120
         Picture         =   "FESTSU07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Cierra y Abandona la Aplicación"
         Top             =   240
         Width           =   800
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Help"
         Height          =   700
         Left            =   2280
         Picture         =   "FESTSU07.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   240
         Width           =   800
      End
      Begin VB.CommandButton Command2 
         Height          =   700
         Left            =   120
         Picture         =   "FESTSU07.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Re-Calcula Resumen Estadístico de Consumo"
         Top             =   240
         Width           =   800
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   1200
         ScaleHeight     =   75
         ScaleWidth      =   795
         TabIndex        =   37
         Top             =   120
         Width           =   855
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   38
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Imprimir"
         Enabled         =   0   'False
         Height          =   700
         Left            =   120
         Picture         =   "FESTSU07.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Aprueba - Graba e Imprime Remito"
         Top             =   6960
         Width           =   800
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Trasla."
         Height          =   700
         Left            =   120
         Picture         =   "FESTSU07.frx":0DC8
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Generar Diskette de Transferencia"
         Top             =   4920
         Visible         =   0   'False
         Width           =   800
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   150
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   19
         Top             =   6000
         Width           =   650
         Begin VB.Frame Frame10 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   20
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   510
         Left            =   620
         Picture         =   "FESTSU07.frx":10D2
         ToolTipText     =   "Doble Click para Configurar Modalidad de Cálculo"
         Top             =   420
         Width           =   2010
      End
      Begin VB.Image Image3 
         Height          =   390
         Left            =   -250
         Picture         =   "FESTSU07.frx":46BC
         Top             =   6240
         Width           =   1515
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0EEC0&
      Caption         =   "Listados de Activación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   2520
      TabIndex        =   11
      Top             =   6300
      Width           =   3900
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   945
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   12
         Top             =   945
         Width           =   2430
         Begin VB.CommandButton Command4 
            Height          =   600
            Left            =   -20
            Picture         =   "FESTSU07.frx":65DE
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Activación de Recepción"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Recepciones Pen- dientes de Aprobar"
            Height          =   750
            Left            =   630
            TabIndex        =   14
            Top             =   75
            Width           =   1905
         End
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   525
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   15
         Top             =   420
         Width           =   2430
         Begin VB.CommandButton Command5 
            Height          =   600
            Left            =   0
            Picture         =   "FESTSU07.frx":6EA8
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Activación Compras Pendientes"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Listado de Activa- ción de Compras"
            Height          =   540
            Left            =   675
            TabIndex        =   17
            Top             =   45
            Width           =   1710
         End
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0EEC0&
      Caption         =   "Análisis Gráfico"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   5145
      TabIndex        =   7
      Top             =   1470
      Width           =   3900
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   525
         ScaleHeight     =   585
         ScaleWidth      =   2790
         TabIndex        =   8
         Top             =   500
         Width           =   2850
         Begin VB.CommandButton Command17 
            Height          =   600
            Left            =   0
            Picture         =   "FESTSU07.frx":7772
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Gráfico Estadístico de Consumo Mensual por Pieza"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Estadísticas de Consumo por Artículo"
            Height          =   750
            Left            =   735
            TabIndex        =   10
            Top             =   105
            Width           =   2010
         End
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0EEC0&
      Caption         =   "Consultas y Listados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      Left            =   5145
      TabIndex        =   0
      Top             =   2940
      Width           =   3900
      Begin VB.PictureBox Picture7 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   840
         ScaleHeight     =   585
         ScaleWidth      =   2475
         TabIndex        =   1
         Top             =   945
         Width           =   2535
         Begin VB.CommandButton Command15 
            Height          =   600
            Left            =   -20
            Picture         =   "FESTSU07.frx":7A7C
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Listado Estadístico de Entregas"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Análisis Estadístico Entrega / Sucursal"
            Height          =   750
            Left            =   630
            TabIndex        =   3
            Top             =   80
            Width           =   1905
         End
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   525
         ScaleHeight     =   585
         ScaleWidth      =   2475
         TabIndex        =   4
         Top             =   420
         Width           =   2535
         Begin VB.CommandButton Command19 
            Height          =   600
            Left            =   0
            Picture         =   "FESTSU07.frx":8776
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Listado de Recomendación de Reposición"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Reposición  Reco- mendada por Pieza"
            Height          =   540
            Left            =   675
            TabIndex        =   6
            Top             =   45
            Width           =   1815
         End
      End
   End
   Begin VB.Line Line1 
      X1              =   5985
      X2              =   7245
      Y1              =   1200
      Y2              =   1200
   End
End
Attribute VB_Name = "FESTSU07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DESDEFE%, HASTAFE%
Dim DESDEMES%, HASTAMES%
Dim CAMESES%, DESMES$, HASMES$
'
Dim VEDEP$, VECLIEN$
'
'Dim CORRESP$, ABREMES$, CORRESQ$
'Dim DESFE$, HASFE$
'Dim CODIMES%
'
'Dim GRCC%(32767)
'Dim GRCV%(32767)
'Dim COREGIS#(32767)
'
'Dim MONECO%      ' 1=Pesos, 2=Dolar
'
Private Sub Command15_Click()
   Command15.Enabled = False
   Call HEADERS("STATSUCU", "")
   CLISUC$ = "VARIOS"
   If List1.ListCount = 1 Then
     CLISUC$ = List1.List(0)
   End If
   CLISUC$ = CLISUC$ + " / " + TRIM$(Label11)
   Call HEADERS("STATSUCU", "Período: " + TRIM$(label1) + " - " + TRIM$(Label9) + " (" + TRIM$(Label12) + " Meses)")
   Call HEADERS("STATSUCU", "Cliente / Sucursal: " + CLISUC$)
   '
   Call FINAL("*ANASTSUC")
   Command15.Enabled = True
End Sub

Private Sub Command17_Click()
   Command17.Enabled = False
   ANESTS07.Show 1
   Command17.Enabled = True
End Sub

Private Sub Command19_Click()
   Command19.Enabled = False
   RETECO$ = ANESTS07.MESAN
   ATENTI% = 1
   '
   FIN& = ULTREG&("STATSUCU")
   Screen.MousePointer = 1
   JJ& = 0
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XX$ = REGLEIDO$("STATSUCU", U&)
     RESTAT$ = XX$
     ANESTS07.MESAN = Left$(XX$, 15)
     TTXX$ = REGBLAN$("CRITSTO")
     CIXI% = CODINT%(Left$(XX$, 15))
     STOAC = CVS(Mid$(XX$, 77, 4))
     MINIREC = Val(ANESTS07.Text1)
     LREPREC = Val(ANESTS07.Text3)
     '
     Call REPLA(TTXX$, MKI$(CIXI%), 1, 2)
     Call REPLA(TTXX$, UNIMED$(CIXI%), 3, 2)
     Call REPLA(TTXX$, MKS$(STOAC), 5, 4)
     Call REPLA(TTXX$, MKS$(MINIREC), 9, 4)
       FALTA = 0: If STOAC < MINIREC Then FALTA = MINIREC - STOAC
     Call REPLA(TTXX$, MKS$(FALTA), 13, 4)
       If FALTA < 0.05 Then LREPREC = 0
     Call REPLA(TTXX$, MKS$(LREPREC), 17, 4)
     '
     JJ& = JJ& + 1
     Call GRAREG("CRITSTO", TTXX$, JJ&)
   Next U&
   Call SETULTREG("CRITSTO", JJ&)
   ATENTI% = 0
   ANESTS07.MESAN = RETECO$
   '
   Call HEADERS("CRITSTO", "")
   CLISUC$ = "VARIOS"
   If List1.ListCount = 1 Then
     CLISUC$ = List1.List(0)
   End If
   CLISUC$ = CLISUC$ + " / " + TRIM$(Label11)
   Call HEADERS("CRITSTO", "Período: " + TRIM$(label1) + " - " + TRIM$(Label9) + " (" + TRIM$(Label12) + " Meses)")
   Call HEADERS("CRITSTO", "Cliente / Sucursal: " + CLISUC$)
   '
   Call FINAL("*REPOREC")
   Command19.Enabled = True
End Sub


Private Sub Command4_Click()
   Command4.Enabled = False
   Screen.MousePointer = 11
   '
   JJ& = 0: KK& = 0
   For U& = 1 To ULTREG&("BOLREPEN")
     X$ = REGLEIDO$("BOLREPEN", U&)
     NBR& = CVS(Left$(X$, 4))
     If NBR& > 0 Then
       KK& = KK& + 1
       Call GRAREG("BOLREPEN", X$, KK&)
       '
       APRO# = CVD(Mid$(X$, 163, 8))
       RECHA# = CVD(Mid$(X$, 171, 8))
       FALTA# = CVD(Mid$(X$, 179, 8))
       If Abs(APRO#) < 0.005 Then
         If Abs(RECHA#) < 0.005 Then
           If Abs(FALTA#) < 0.005 Then
             Call REPLA(X$, String$(128, 0), 155, 128)
             FEBOX = CVI(Mid$(X$, 81, 2))
             SATRA = (DIENTRE(FEBOX, HOY(HOS$))) / 7
             Call REPLA(X$, MKS$(SATRA), 155, 4)
             '
             JJ& = JJ& + 1
             Call GRAREG("ACTIREPA", X$, JJ&)
           End If
         End If
       End If
     End If
   Next U&
   '
   Call SETULTREG("BOLREPEN", KK&)
   Call SETULTREG("ACTIREPA", JJ&)
   Call CIERRARCH("IBOLPEN")
   '
   Call FINAL("*ACTIREPA")
   '
   Screen.MousePointer = 1
   Command4.Enabled = True
   '
End Sub

Private Sub Command5_Click()
   '
   Command5.Enabled = False
   '
   Screen.MousePointer = 11
   HOI1 = HOY(HOS$)
   FIN& = ULTREG&("OCOMDETA")
   JJ& = 0
   '
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XXX$ = REGLEIDO$("OCOMDETA", U&)
     If Asc(Mid$(XXX$, 124, 1)) < 8 Then ' ordenes abiertas
       If CVS(Mid$(XXX$, 99, 4)) < CVS(Mid$(XXX$, 55, 4)) - 0.1 Then
         XXY$ = XXX$ + String$(32, 0)
         CAPEN = CVS(Mid$(XXX$, 55, 4)) - CVS(Mid$(XXX$, 99, 4))
         If CAPEN < 0.05 Then CAPEN = 0
         Call REPLA(XXY$, MKS$(CAPEN), 129, 4)
         FENTRESO = CVI(Mid$(XXX$, 95, 2))
         If FENTRESO <= HOI1 Then
           DIATRA = DIENTRE%(FENTRESO, HOI1)
           Call REPLA(XXY$, MKS$(DIATRA / 7), 149, 4)
           JJ& = JJ& + 1
           Call GRAREG("OPENDETA", XXY$, JJ&)
         End If
       End If
     End If
   Next U&
   '
   Call SETULTREG("OPENDETA", JJ&)
   Call CIERRARCH("OPENDETA")
   '
   Call FILESORT("OPENDETA", "", 3, 3, "ASC")
   Call FINAL("*LACTICOM")
   '
   Screen.MousePointer = 1
   '
   Command5.Enabled = True
End Sub

Private Sub Form_Load()
    '
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    If EXISTE%(LUDAT$ + "STATSUCU.DAT") > 0 Then
      TTXX$ = Format(FileDateTime(LUDAT$ + "STATSUCU.DAT"), "dd/mm/yy hh:mm")
      FEX = FECHANUM(Left$(TTXX$, 8))
      Label3 = DSEMIN$(FEX) + " " + Left$(TTXX$, 8) + " - " + Mid$(TTXX$, 10) + " hs"
      TTXX$ = CONTROL$("ESTACONS", "PERICALC")
      label1 = TRIM$(Left$(TTXX$, 16))
      Label9 = TRIM$(Mid$(TTXX$, 17, 16))
      Label12 = TRIM$(Mid$(TTXX$, 33, 2))
      If Val(Label12) < 6 Or Val(Label12) > 24 Or (Val(Label12) Mod 2) <> 0 Then Label12 = ""
      Label11 = CONTROL$("ESTACONS", "DEPOSEL")
      '
      List1.Clear
      TECLIENTES$ = LOADFILE$(LUDAT$ + "CLISELST.SLX")
10    PPXX% = InStr(TECLIENTES$, Chr$(124))
      If PPXX% > 0 Then
        CLIENX$ = Left$(TECLIENTES$, PPXX% - 1)
        If CLIENX$ <> "" Then
          List1.AddItem CLIENX$
        End If
        TECLIENTES$ = Mid$(TECLIENTES$, PPXX% + 1)
        GoTo 10
      End If
    End If
    '
    Show
    Refresh
    Call ABRECONEXION
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Command1_Click()
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Image2_DblClick()
   OPCALCOS.Show 1
End Sub

Private Sub Command2_Click()
   '
   Command2.Enabled = False
   OPSTATSU07.label1 = "8"
10 OPSTATSU07.Show 1
   DoEvents
   If Val(OPSTATSU07.label1) < 1 Then GoTo 99
   '
   VEDEP$ = String$(256, 0)
   VECLIEN$ = String$(32767, 0)
   '
   If OPSTATSU07.List1.ListCount < 1 Then
     Call MENSERR(24, "Falta seleccionar Clientes")
     GoTo 10
   End If
   '
   List1.Clear
   TECLIENTES$ = ""
   For U& = 1 To OPSTATSU07.List1.ListCount
     NCLIEL& = Val(Mid$(OPSTATSU07.List1.List(U& - 1), 65, 8))
     If NCLIEL& > 0 Then
       If NCLIEL& <= 32767 Then
         Mid$(VECLIEN$, NCLIEL&, 1) = Chr$(1)
         List1.AddItem OPSTATSU07.List1.List(U& - 1)
       End If
     End If
   Next U&
   '
   For U& = 1 To OPSTATSU07.List3.ListCount
     If OPSTATSU07.List3.Selected(U& - 1) = True Then GoTo 20
   Next U&
   Call MENSERR(24, "Debe Seleccionar al Menos una Sucursal")
   GoTo 10
   '
20 VEDEPOS$ = ""
   For U& = 1 To OPSTATSU07.List3.ListCount
     If OPSTATSU07.List3.Selected(U& - 1) = True Then
       If VEDEPOS$ <> "" Then VEDEPOS$ = VEDEPOS$ + " + "
       VEDEPOS$ = VEDEPOS$ + UCase$(TRIM$(Mid$(OPSTATSU07.List3.List(U& - 1), 65, 8)))
       NDEP& = Val(Left$(OPSTATSU07.List3.List(U& - 1), 3))
       If NDEP& > 0 Then
         If NDEP& <= 256 Then
           Mid$(VEDEP$, NDEP&, 1) = Chr$(1)
         End If
       End If
     End If
   Next U&
   Label11 = VEDEPOS$
   Label11.Refresh
   '
30 Call ARMAMES
   
   Call CALCURESU
   '
99 Command2.Enabled = True
   '
End Sub

Sub ARMAMES()
   '
   DESDEFE% = 0: HASTAFE% = 0
   CORRESP$ = TRIM$(OPSTATSU07.MESAN)
   CORRESQ$ = MESANO$(CORRESP$, DESDEFE%, HASTAFE%)
   '
   If DESDEFE% < 1 Then
     Call MENSERR(24, "Falta Indicar Período Mensual")
   End If
   '
   FEX = DESDEFE%: DESFE$ = FECHATEX$(FEX)
   FEX = HASTAFE%: HASFE$ = FECHATEX$(FEX)
   '
   CORRESQ$ = TRIM$(CORRESP$)
   ABREMES$ = Left$(CORRESQ$, 3) + Mid$(DESFE$, 7, 2)
   ANIO% = Val(Mid$(DESFE$, 7, 2)): If ANIO% < 80 Then ANIO% = 100 + ANIO%
   MESO% = Val(Mid$(DESFE$, 4, 2))
   HASTAMES% = 12 * ANIO% + MESO%: HASMES$ = ABREMESAN$(HASTAMES%, 1)
   '
   CAMESES% = Val(OPSTATSU07.label1)
   '
   For UK% = 1 To CAMESES% - 1
     MESO% = MESO% - 1
     If MESO% < 1 Then
       MESO% = MESO% + 12
       ANIO% = ANIO% - 1
     End If
   Next UK%
   DESDEMES% = 12 * ANIO% + MESO%: DESMES$ = ABREMESAN$(DESDEMES%, 1)
   MESOX$ = Right$("00" + TRIM$(Str$(MESO%)), 2)
   ANIOX$ = Right$("00" + TRIM$(Str$(ANIO% Mod 100)), 2)
   DESDEFE% = FECHANUM("01/" + MESOX$ + "/" + ANIOX$)
   '
End Sub

Sub CALCURESU()
    '
   Call ARMAMES
   If DESDEFE% < 1 Then Exit Sub
   '
   Screen.MousePointer = 11
   Dim TOPED(23, 32767)
   Dim TOCON(23, 32767)
   Dim ACCLI%(32767)
   Dim IVEN(2, 48)
   Dim ACUCO%(2)
   '
   DESDE = FETEXCOR1(DESDEFE%)
   HASTA = FETEXCOR1$(HASTAFE%)
   '
   Dim RS As ADODB.Recordset
   SQLX$ = "SELECT * FROM MOVISTO"
   SQLX$ = SQLX$ + " WHERE FECHMOV >='" & DESDE & "' AND FECHMOV <='" & HASTA & "'"
   Set RS = FLEXCONN.Execute(SQLX$)
      
   CONDI$ = "FECHMOV >='" & DESDE & "' AND FECHMOV <='" & HASTA & "'"
   FIN& = CantRegistros("MOVISTO", CONDI$)
   With RS
        Do While Not .EOF
            Call TRACE(20, U&, FIN&)
            NDEP% = XVALO(!DepoMovi)
            If NDEP% < 1 Then NDEP% = 1
            If NDEP% <= 256 Then
              If Asc(Mid$(VEDEP$, NDEP%, 1)) = 1 Then
                NCLIE% = XVALO(!Nume_Clie)
                If NCLIE% > 0 Or NDEP% = 10 Then  ' PARA PREVER CASO NEUQUEN
                  If NCLIE% <= 32767 Then
                    If NDEP% = 10 Then GoTo 15    ' PARA PREVER CASO NEUQUEN
                    If Asc(Mid$(VECLIEN$, NCLIE%, 1)) = 1 Then
15                    CAPED = XVALO(!Cantsalid)
                      If CAPED > 0.01 Then
                        'If Mid$(X$, 23, 2) = "OF" Then
                          FENTRESOL% = CVINT(!fechMov)
                          FEX = FENTRESOL%: FESOL$ = FECHATEX$(FEX)
                          ANIOX% = Val(Mid$(FESOL$, 7, 2)): If ANIOX% < 80 Then ANIOX% = 100 + ANIOX%
                          CODIPE% = 12 * ANIOX% + Val(Mid$(FESOL$, 4, 2))
                          CI% = XVALO(!Cod_Inte)
                          If CI% > 0 And CI% <= NUMAS% Then
                            TOCON(CODIPE% - DESDEMES%, CI%) = TOCON(CODIPE% - DESDEMES%, CI%) + CAPED
                            ACCLI%(CI%) = 1
                          End If
                        'End If
                      End If
                    End If
                  End If
                End If
              End If
            End If
            U& = U& + 1
            .MoveNext
        Loop
   End With
   '
   Screen.MousePointer = 1
   '
   Call GENECOMA
   '
   Screen.MousePointer = 11
   REBLA$ = REGBLAN$("STATSUCU")
   JJ& = 0
   FIN& = NUINV%
   For U& = 1 To NUINV%
     Call TRACE(20, U&, FIN&)
     X$ = REGLEIDO$("INVERT", U& + 1)
     CI% = CVI(Mid$(X$, 17, 2))
     If CI% > 0 Then
       If ACCLI%(CI%) > 0 Then
         Y$ = REBLA$
         Call REPLA(Y$, Left$(X$, 16), 1, 16)
         Call REPLA(Y$, DESCRIT0$(CI%), 17, 48)
         Y0$ = "": Z0$ = ""
         For UKI% = 0 To 23
           Y0$ = Y0$ + MKS$(TOPED(UKI%, CI%))
           Z0$ = Z0$ + MKS$(TOCON(UKI%, CI%))
         Next UKI%
         '
         Y1$ = "": Z1$ = ""
         SUCOVEN = 0: SUCOFAB = 0
         SUVENI = 0: SUVEND = 0
         SUFABI = 0: SUFABD = 0
         SUTOTI = 0: SUTOTD = 0
         MAXIVEN = 0: MAXIPRO = 0: MAXITOT = 0
         MINIVEN = 9999999: MINIFAB = 9999999: MINITOT = 9999999
         For UJ% = 1 To CAMESES%
           '
           COVEN = CVS(Mid$(Y0$, 4 * UJ% - 3, 4))
           SUCOVEN = SUCOVEN + COVEN
           If UJ% <= CAMESES% / 2 Then SUVENI = SUVENI + COVEN Else SUVEND = SUVEND + COVEN
           If COVEN > MAXIVEN Then MAXIVEN = COVEN
           If COVEN < MINIVEN Then MINIVEN = COVEN
           '
           COFAB = CVS(Mid$(Z0$, 4 * UJ% - 3, 4))
           SUCOFAB = SUCOFAB + COFAB
           If UJ% <= CAMESES% / 2 Then SUFABI = SUFABI + COFAB Else SUFABD = SUFABD + COFAB
           If COFAB > MAXIPRO Then MAXIPRO = COFAB
           If COFAB < MINIFAB Then MINIFAB = COFAB
           '
           If COVEN + COFAB > MAXITOT Then MAXITOT = COVEN + COFAB
           If COVEN + COFAB < MINITOT Then MINITOT = COVEN + COFAB
           If UJ% <= CAMESES% / 2 Then SUTOTI = SUTOTI + COVEN + COFAB Else SUTOTD = SUTOTD + COVEN + COFAB
         Next UJ%
         '
         ESTIM = Int((4 * SUVEND - 2 * SUVENI) / CAMESES% + 0.5)
         If ESTIM < 0 Then ESTIM = 0
         If ESTIM < SUCOVEN / CAMESES% Then ESTIM = SUCOVEN / CAMESES%
         ESTIVE = ESTIM
         Y1$ = MKS$(MINIVEN) + MKS$(MAXIVEN) + MKS$(SUCOVEN / CAMESES%) + MKS$(ESTIM)
         '
         ESTIM = Int((4 * SUFABD - 2 * SUFABI) / CAMESES% + 0.5)
         If ESTIM < 0 Then ESTIM = 0
         If ESTIM < SUCOFAB / CAMESES% Then ESTIM = SUCOFAB / CAMESES%
         ESTIFA = ESTIM
         Z1$ = MKS$(MINIFAB) + MKS$(MAXIPRO) + MKS$(SUCOFAB / CAMESES%) + MKS$(ESTIM)
         '
         ESTIM = 0
         For KKI% = 1 To 255
           If Asc(Mid$(VEDEP$, KKI%, 1)) > 0 Then
             ESTIM = ESTIM + STODEPOS(CI%, KKI%)
           End If
         Next KKI%
         If ESTIM < 0 Then ESTIM = 0
         T1$ = MKS$(MINITOT) + MKS$(MAXITOT) + MKS$((SUCOVEN + SUCOFAB) / CAMESES%) + MKS$(ESTIM)
         '
         Call REPLA(Y$, T1$, 65, 16)
         Call REPLA(Y$, Y0$, 129, 96)
         Call REPLA(Y$, Y1$, 225, 16)
         Call REPLA(Y$, Z0$, 257, 96)
         Call REPLA(Y$, Z1$, 353, 16)
         JJ& = JJ& + 1
         Call GRAREG("STATSUCU", Y$, JJ&)
       End If
     End If
   Next U&
   '
   Call SETULTREG("STATSUCU", JJ&)
   '
   'Call COCOBREC
   '
   TTXX$ = Space$(40)
   Call REPLA(TTXX$, DESMES$, 1, 16)
   Call REPLA(TTXX$, HASMES$, 17, 16)
   Call REPLA(TTXX$, TRIM$(Str$(CAMESES%)), 33, 2)
   Call GRACONTROL("ESTACONS", "PERICALC", TTXX$)
   Call GRACONTROL("ESTACONS", "DEPOSEL", TRIM$(Label11))
   '
   TECLIENTES$ = ""
   For U& = 1 To List1.ListCount
     TECLIENTES$ = TECLIENTES$ + List1.List(U& - 1) + Chr$(124)
   Next U&
   Call SAVEFILE(LUDAT$ + "CLISELST.SLX", TECLIENTES$)
   '
   TTXX$ = Format(FileDateTime(LUDAT$ + "STATSUCU.DAT"), "dd/mm/yy hh:mm")
   FEX = FECHANUM(Left$(TTXX$, 8))
   Label3 = DSEMIN$(FEX) + " " + Left$(TTXX$, 8) + " - " + Mid$(TTXX$, 10) + " hs"
   label1 = TRIM$(DESMES$)
   Label9 = TRIM$(HASMES$)
   Label12 = TRIM$(Str$(CAMESES%))
   '
90 Screen.MousePointer = 1
   '
End Sub


Sub GENECOMA()

   Screen.MousePointer = 11
   '
   XX$ = REGBLAN$("MASTER")
   Call REPLA(XX$, MKI$(NUMAS%), 1, 2)
   Call GRAREG("MASTER", XX$, 1)
   REG& = 1
   Sql$ = "SELECT * FROM MASTER order by codint asc"
   Set RS = FLEXCONN.Execute(Sql$)
   With RS
        Do While Not .EOF
            CI% = XVALO(!CODINT)
            DENOMINACION$ = SAFETEXT$(!Descripcion)
            GRUIVA% = XVALO(!GRUPIVA)
            CIEXT$ = SAFETEXT$(!CODIGO)
            TIPOMATE% = XVALO(!TIMATE)
            CRITREPOS% = XVALO(!CRITREP)
            Monecos% = XVALO(!Monecos)
            TCOLOR$ = SAFETEXT$(!COLOR)
            TTALLE$ = SAFETEXT$(!TALLE)
                FECHA$ = FECHATEX$(CVINT(!FEREVIS))
            CCOMPO% = Val(FECHA$)
            PESU = XVALO(!PESUNI)
            LOTEMAI% = XVALO(!STOMAX)
            STOCKMIN# = XVALO(!STOMIN)
            LOTEREP% = XVALO(!LOREPOS)
            Unidad$ = SAFETEXT$(!UMEDIDA)
            COSTO# = XVALO(!COSUNI)
            FECHACOS% = CVINT(!FECOSTO)
            Revision$ = SAFETEXT$(!Revision)
            FECHREV% = CVINT(!FEREVIS)
            PROBHAB% = XVALO(!PROVHABI)
            GRUPOCON% = XVALO(!Grucoven)
            NACIONA% = XVALO(!PORNACIO)
            STAT% = XVALO(!Status)
            '
            X$ = REGBLAN$("MASTER")
            Call REPLA(XX$, MKI$(CI%), 1, 2)
            Call REPLA(XX$, Chr$(GRUIVA%), 3, 1)
            Call REPLA(XX$, DENOMINACION$, 5, 64)
            Call REPLA(XX$, Space$(8), 69, 8)
            Call REPLA(XX$, CIEXT$, 77, 16)
            If STAT% < 0 Then Call REPLA(XX$, Chr$(96), 92, 1)
            Call REPLA(XX$, Chr$(TIPOMATE%), 93, 1)
            Call REPLA(XX$, Chr$(CRITREPOS%), 94, 1)
            Call REPLA(XX$, Chr$(Monecos%), 95, 1)
            Call REPLA(XX$, Chr$(CCOLOR%(TCOLOR$)), 96, 1)
            Call REPLA(XX$, Chr$(CTALLE%(TTALLE$)), 97, 1)
            Call REPLA(XX$, Chr$(CCOMPO%), 98, 1)
            Call REPLA(XX$, MKS$(PESU), 99, 4)
            Call REPLA(XX$, MKI$(LOTEMAI%), 103, 2)
            Call REPLA(XX$, MKI$(STOCKMIN#), 105, 2)
            Call REPLA(XX$, MKI$(LOTEREP%), 107, 2)
            Call REPLA(XX$, Unidad$, 109, 2)
            Call REPLA(XX$, MKS$(COSTO#), 111, 4)
            Call REPLA(XX$, MKI$(FECHACOS%), 115, 2)
            Call REPLA(XX$, Revision$, 117, 4)
            Call REPLA(XX$, MKI$(FECHREV%), 121, 2)
            Call REPLA(XX$, MKI$(PROBHAB%), 123, 2)
            Call REPLA(XX$, MKI$(GRUPOCON%), 125, 2)
            Call REPLA(XX$, MKI$(NACIONA%), 127, 2)
            REG& = REG& + 1
            Call GRAREG("MASTER", XX$, REG&)
            '
            .MoveNext
        Loop
        Call SETULTREG("MASTER", REG&)
        Call CIERRARCH("MASTER")
   End With
   '
   FIN& = ULTREG&("MASTER")
   If FIN& < 2 Then
     Screen.MousePointer = 1
     Exit Sub
   End If
   '
   JI& = 1  ' PARA INVERT1
   REBLI$ = Space$(16) + MKI$(0)
   JE& = 0  ' PARA ECOMAST1
   REBLE$ = Space$(62) + MKI$(0)
   '
   XI$ = REGLEIDO$("INVERT", 1)
   Call REPLA(XI$, String$(16, 0), 1, 16)
   Call GRAREG("INVERT", XI$, 1)
   '
   For U& = 2 To FIN&
     Call TRACE(20, U&, FIN&)
     XM$ = REGLEIDO$("MASTER", U&)
     CII% = U& - 1
     If CVI(Left$(XM$, 2)) <> CII% Then
       Call REPLA(XM$, MKI$(CII%), 1, 2)
       Call GRAREG("MASTER", XM$, U&)
     End If
     '
     CODE$ = Mid$(XM$, 77, 16)
     DEDE$ = Mid$(XM$, 5, 64)
     '
     XI$ = REBLI$
     Call REPLA(XI$, CODE$, 1, 16)
     Call REPLA(XI$, MKI$(CII%), 17, 2)
     JI& = JI& + 1
     Call GRAREG("INVERT", XI$, JI&)
     '
     If Mid$(CODE$, 16, 1) <> Chr$(96) Then
       XE$ = REBLE$
       Call REPLA(XE$, CODE$, 1, 16)
       Call REPLA(XE$, DEDE$, 17, 46)
       Call REPLA(XE$, MKI$(CII%), 63, 2)
       JE& = JE& + 1
       Call GRAREG("ECOMAST1", XE$, JE&)
       '
     End If
   Next U&
   '
   NMS% = FIN& - 1
   XM$ = REGLEIDO$("MASTER", 1)
   Call REPLA(XM$, MKI$(NMS%), 1, 2)
   Call GRAREG("MASTER", XM$, 1)
   Call CIERRARCH("MASTER")
   '
   NVI% = JI& - 1
   XI$ = REGLEIDO$("INVERT", 1)
   Call REPLA(XI$, MKI$(NVI%), 17, 2)
   Call GRAREG("INVERT", XI$, 1)
   Call SETULTREG("INVERT", JI&)
   Call CIERRARCH("INVERT")
   Call FILESORT("INVERT", "", 1, 1, "ASC")
   '
   Call SETULTREG("ECOMAST", JE&)
   Call CIERRARCH("ECOMAST")
   Call FILESORT("ECOMAST", "", 1, 1, "ASC")
   '
   BFC$ = ""
   NMSX& = NUMAS%
   LOPUCO& = 16 * NMSX&
   BFC$ = String$(LOPUCO&, 0)
   For U& = 1 To NUMAS%
      X$ = REGLEIDO$("MASTER", U& + 1)
      BFD$ = AJUSTI$(Mid$(X$, 77, 16), 16)
      Mid$(BFC$, 16 * U& - 15, 16) = BFD$
   Next U&
   '
   LOPUCO1& = 2048 * (1 + Int(LOPUCO& / 2048))
   BFC1$ = Left$(BFC$ + String$(4096, 0), LOPUCO1&)
   Call SAVEFILE(LUDAT$ + "PUNCODIG.DAT", BFC1$)

   Screen.MousePointer = 1
End Sub

Sub COCOBREC()
   '
   Dim REOCO&(8192), PUBOL$(8192), CAPEN(8192)
   '
   HAS% = HASTAFE%
   DES% = DESDEFE%
   '
   Screen.MousePointer = 11
   '
   J& = 0: PUCON$ = ""
   FIN& = ULTREG&("OCOMDETA")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("OCOMDETA", U&)
      FF% = CVI(Mid$(X$, 5, 2))        ' fecha de emision
      NP% = CVI(Mid$(X$, 7, 2))        ' proveedor
      STAT% = Asc(Mid$(X$, 124, 1))
      '
      If STAT% <> 9 Then               ' excluye anuladas
        If NP% > 0 Then
          If NP% <> 9999 Then            ' excluye pedidos suministro
            If FF% >= DES% Then
              If FF% <= HAS% Then
                J& = J& + 1
                If J& > 8192 Then
                  Call MENSERR(24, "Demasiados Movimientos\en el Rango de Fechas Elegido.")
                  Exit Sub
                End If
                '
                REOCO&(J&) = U&
                CAPEN(J&) = CVS(Mid$(X$, 55, 4))
                CLACO$ = Mid$(X$, 7, 2) + Left$(X$, 4) + Mid$(X$, 9, 2)
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
   NCOCO& = 0
   For K& = 1 To CAOCO&
      X$ = REGLEIDO$("OCOMDETA", REOCO&(K&))
      NUOCO& = CVS(Left$(X$, 4))
      FEOCO% = CVI(Mid$(X$, 5, 4))
      CIXI% = CVI(Mid$(X$, 9, 2))
      NPRO% = CVI(Mid$(X$, 7, 2))
      '
      Z0$ = REGBLAN$("COPLAENT")
      Call REPLA(Z0$, MKI$(CIXI%), 1, 2)
      Call REPLA(Z0$, Left$(X$, 10), 3, 10)
      '
      FENTREPRO = FEOCO%
      If Len(PUBOL$(K&)) >= 4 Then
         For II& = 1 To Len(PUBOL$(K&)) Step 4
           Z$ = Z0$
           REBREC& = CVS(Mid$(PUBOL$(K&), II&, 4))
           Y$ = REGLEIDO$("BOLRECEP", REBREC&)
           FERECI = CVI(Mid$(Y$, 81, 2))
           DIATRA% = 0
           If FENTREPRO < FERECI Then
             DIATRA% = DIENTRE%(FENTREPRO, FERECI)
           End If
           FERE% = FERECI
           Call REPLA(Z$, MKI$(FERE%), 13, 2)
           Call REPLA(Z$, MKI$(DIATRA%), 15, 2)
           Call REPLA(Z$, MKS$(DIATRA% / 30), 17, 4)
           '
           NCOCO& = NCOCO& + 1
           Call GRAREG("COPLAENT", Z$, NCOCO&)
         Next II&
      End If
      '
   Next K&
   '
   Call SETULTREG("COPLAENT", NCOCO&)
   Call CIERRARCH("COPLAENT")
   '
   Call FILESORT("COPLAENT", "", 4, 4, "ASC")
   Screen.MousePointer = 1
End Sub



