VERSION 5.00
Begin VB.Form FESTOK00 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control Estadístico de Consumos"
   ClientHeight    =   5010
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
   ScaleHeight     =   5010
   ScaleWidth      =   9210
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0E0FF&
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
      Left            =   5145
      TabIndex        =   29
      Top             =   3045
      Width           =   3900
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   945
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   30
         Top             =   945
         Width           =   2430
         Begin VB.CommandButton Command4 
            Height          =   600
            Left            =   -20
            Picture         =   "FESTOK00.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   31
            ToolTipText     =   "Activación de Recepción"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Recepciones Pen- dientes de Aprobar"
            Height          =   750
            Left            =   630
            TabIndex        =   32
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
         TabIndex        =   33
         Top             =   420
         Width           =   2430
         Begin VB.CommandButton Command5 
            Height          =   600
            Left            =   0
            Picture         =   "FESTOK00.frx":08CA
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Activación Compras Pendientes"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Listado de Activa- ción de Compras"
            Height          =   540
            Left            =   675
            TabIndex        =   35
            Top             =   45
            Width           =   1710
         End
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0FF&
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
      TabIndex        =   25
      Top             =   1470
      Width           =   3900
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   525
         ScaleHeight     =   585
         ScaleWidth      =   2790
         TabIndex        =   26
         Top             =   500
         Width           =   2850
         Begin VB.CommandButton Command17 
            Height          =   600
            Left            =   0
            Picture         =   "FESTOK00.frx":1194
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Gráfico Estadístico de Consumo Mensual por Pieza"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Estadísticas de Consumo por Artículo"
            Height          =   750
            Left            =   735
            TabIndex        =   28
            Top             =   105
            Width           =   2010
         End
      End
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   6195
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   23
      Top             =   420
      Width           =   855
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   24
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.CommandButton Command2 
      Height          =   900
      Left            =   5145
      Picture         =   "FESTOK00.frx":149E
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Re-Calcula Resumen Estadístico de Consumo"
      Top             =   315
      Width           =   855
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0FF&
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
      Height          =   2325
      Left            =   210
      TabIndex        =   4
      Top             =   2520
      Width           =   4845
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1470
         ScaleHeight     =   585
         ScaleWidth      =   2685
         TabIndex        =   11
         Top             =   1470
         Width           =   2745
         Begin VB.CommandButton Command18 
            Height          =   600
            Left            =   0
            Picture         =   "FESTOK00.frx":17A8
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Análisis de Disponibilidad"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Acceso a Disponibili- dad General de Stock"
            Height          =   540
            Left            =   735
            TabIndex        =   13
            Top             =   105
            Width           =   1905
         End
      End
      Begin VB.PictureBox Picture7 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1155
         ScaleHeight     =   585
         ScaleWidth      =   2580
         TabIndex        =   5
         Top             =   945
         Width           =   2640
         Begin VB.CommandButton Command15 
            Height          =   600
            Left            =   -20
            Picture         =   "FESTOK00.frx":1AB2
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Historial de Plazoz de Entrega"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Análisis de Plazos de Entrega por Artículo"
            Height          =   750
            Left            =   630
            TabIndex        =   7
            Top             =   45
            Width           =   1905
         End
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   840
         ScaleHeight     =   585
         ScaleWidth      =   2580
         TabIndex        =   8
         Top             =   420
         Width           =   2640
         Begin VB.CommandButton Command19 
            Height          =   600
            Left            =   0
            Picture         =   "FESTOK00.frx":27AC
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Cálculo Stock Mínimo Recomendado"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Stock Mínimo Reco- mendado por Pieza"
            Height          =   540
            Left            =   675
            TabIndex        =   10
            Top             =   45
            Width           =   1815
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   2220
      Left            =   210
      TabIndex        =   2
      Top             =   210
      Width           =   4845
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Meses"
         Height          =   1065
         Left            =   3465
         TabIndex        =   19
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
            TabIndex        =   20
            Top             =   315
            Width           =   675
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Período"
         Height          =   1065
         Left            =   210
         TabIndex        =   14
         Top             =   420
         Width           =   3165
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
            TabIndex        =   18
            Top             =   630
            Width           =   1935
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Hasta:"
            Height          =   225
            Left            =   210
            TabIndex        =   17
            Top             =   690
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
            TabIndex        =   16
            Top             =   315
            Width           =   1935
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Desde:"
            Height          =   225
            Left            =   210
            TabIndex        =   15
            Top             =   380
            Width           =   750
         End
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "F.Cálculo:"
         Height          =   225
         Index           =   0
         Left            =   105
         TabIndex        =   21
         Top             =   1780
         Width           =   1065
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
         TabIndex        =   3
         Top             =   1725
         Width           =   3130
      End
   End
   Begin VB.CommandButton command1 
      Caption         =   "Exit"
      Height          =   900
      Left            =   8190
      Picture         =   "FESTOK00.frx":2AB6
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Cierra y Abandona la Aplicación"
      Top             =   315
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Help"
      Height          =   900
      Left            =   7245
      Picture         =   "FESTOK00.frx":2C00
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
      Top             =   315
      Width           =   855
   End
   Begin VB.Line Line1 
      X1              =   5985
      X2              =   7245
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   5670
      Picture         =   "FESTOK00.frx":2F0A
      ToolTipText     =   "Doble Click para Configurar Modalidad de Cálculo"
      Top             =   705
      Width           =   2010
   End
End
Attribute VB_Name = "FESTOK00"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DESDEFE%, HASTAFE%
Dim DESDEMES%, HASTAMES%
Dim CAMESES%, DESMES$, HASMES$
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
   Call COPYSTRU("COPLAENT", "LIPLAENT")
   Call FILESORT("COPLAENT", "LIPLAENT", 1, 3, "ASC")
   Call FINAL("*LIPLAENT")
   Command15.Enabled = True
End Sub

Private Sub Command17_Click()
   Command17.Enabled = False
   ANESTOCK.Show 1
   Command17.Enabled = True
End Sub

Private Sub Command18_Click()
   Command18.Enabled = False
   DISPOSTO.Show 1
   Command18.Enabled = True
End Sub

Private Sub Command19_Click()
   Command19.Enabled = False
   RETECO$ = ANESTOCK.MESAN
   ATENTI% = 1
   FIN& = ULTREG&("STATSTOK")
   Screen.MousePointer = 1
   JJ& = 0
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XX$ = REGLEIDO$("STATSTOK", U&)
     RESTAT$ = XX$
     ANESTOCK.MESAN = Left$(XX$, 15)
     TTXX$ = REGBLAN$("STMINREC")
     CIXI% = CODINT%(ANESTOCK.MESAN)
     MINIREC = Val(ANESTOCK.Text1)
     LREPREC = Val(ANESTOCK.Text2)
     '
     Call REPLA(TTXX$, MKI$(CIXI%), 1, 2)
     Call REPLA(TTXX$, UNIMED$(CIXI%), 3, 2)
     Call REPLA(TTXX$, MKS$(Val(ANESTOCK.Label13)), 5, 4)
     Call REPLA(TTXX$, MKS$(MINIREC), 9, 4)
     Call REPLA(TTXX$, MKS$(Val(ANESTOCK.Label16)), 13, 4)
     Call REPLA(TTXX$, MKS$(LREPREC), 17, 4)
     '
     If CVS(Mid$(TTXX$, 9, 4)) > CVS(Mid$(TTXX$, 5, 5)) + 0.05 Then
       Call REPLA(TTXX$, "Ajustar St.Minimo", 21, 28)
     End If
     '
     JJ& = JJ& + 1
     Call GRAREG("STMINREC", TTXX$, JJ&)
   Next U&
   Call SETULTREG("STMINREC", JJ&)
   ATENTI% = 0
   ANESTOCK.MESAN = RETECO$
   '
   Call FINAL("*STMINREC")
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
         xxy$ = XXX$ + String$(32, 0)
         CAPEN = CVS(Mid$(XXX$, 55, 4)) - CVS(Mid$(XXX$, 99, 4))
         If CAPEN < 0.05 Then CAPEN = 0
         Call REPLA(xxy$, MKS$(CAPEN), 129, 4)
         FENTRESO = CVI(Mid$(XXX$, 95, 2))
         If FENTRESO <= HOI1 Then
           DIATRA = DIENTRE%(FENTRESO, HOI1)
           Call REPLA(xxy$, MKS$(DIATRA / 7), 149, 4)
           JJ& = JJ& + 1
           Call GRAREG("OPENDETA", xxy$, JJ&)
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
    If EXISTE%(LUDAT$ + "STATSTOK.DAT") > 0 Then
      TTXX$ = Format(FileDateTime(LUDAT$ + "STATSTOK.DAT"), "dd/mm/yy hh:mm")
      FEX = FECHANUM(Left$(TTXX$, 8))
      Label3 = DSEMIN$(FEX) + " " + Left$(TTXX$, 8) + " - " + Mid$(TTXX$, 10) + " hs"
      TTXX$ = CONTROL$("ESTACONS", "PERICALC")
      label1 = TRIM$(Left$(TTXX$, 16))
      Label9 = TRIM$(Mid$(TTXX$, 17, 16))
      Label12 = TRIM$(Mid$(TTXX$, 33, 2))
      If Val(Label12) < 6 Or Val(Label12) > 24 Or (Val(Label12) Mod 2) <> 0 Then Label12 = ""
    End If
    '
    Show
    Refresh
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
   OPCALSTAT.label1 = "8"
   OPCALSTAT.Show 1
   DoEvents
   If Val(OPCALSTAT.label1) < 1 Then GoTo 99
   '
   Call ARMAMES
   Call CALCURESU
   '
99 Command2.Enabled = True
   '
End Sub

Sub ARMAMES()
   '
   DESDEFE% = 0: HASTAFE% = 0
   CORRESP$ = TRIM$(OPCALSTAT.MESAN)
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
   CAMESES% = Val(OPCALSTAT.label1)
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
   FIN& = ULTREG&("PEDDETA")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     X$ = REGLEIDO$("PEDDETA", U&)
     '        NC% = CVI(Mid$(X$, 7, 2))
     If Asc(Mid$(X$, 27, 1)) <> 9 Then ' no esta anulado
       FENTRESOL% = CVI(Mid$(X$, 25, 2))
       If FENTRESOL% >= DESDEFE% Then
         If FENTRESOL% <= HASTAFE% Then
           FEX = FENTRESOL%: FESOL$ = FECHATEX$(FEX)
           ANIOX% = Val(Mid$(FESOL$, 7, 2)): If ANIOX% < 80 Then ANIOX% = 100 + ANIOX%
           CODIPE% = 12 * ANIOX% + Val(Mid$(FESOL$, 4, 2))
           CI% = CVI(Mid$(X$, 33, 2))
           If CI% > 0 And CI% <= NUMAS% Then
             CAPED = CVS(Mid$(X$, 79, 4))
             TOPED(CODIPE% - DESDEMES%, CI%) = TOPED(CODIPE% - DESDEMES%, CI%) + CAPED
             ACCLI%(CI%) = 1
           End If
         End If
       End If
     End If
   Next U&
   Call CIERRARCH("PEDDETA")
   '
   FIN& = ULTREG&("MOVISTO")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     X$ = REGLEIDO$("MOVISTO", U&)
     '        NC% = CVI(Mid$(X$, 7, 2))
     FENTRESOL% = CVI(Mid$(X$, 1, 2))
     If FENTRESOL% >= DESDEFE% Then
       If FENTRESOL% <= HASTAFE% Then
         CAPED = CVD(Mid$(X$, 96, 8))
         If CAPED > 0.01 Then
           If Mid$(X$, 23, 2) = "OF" Then
             FEX = FENTRESOL%: FESOL$ = FECHATEX$(FEX)
             ANIOX% = Val(Mid$(FESOL$, 7, 2)): If ANIOX% < 80 Then ANIOX% = 100 + ANIOX%
             CODIPE% = 12 * ANIOX% + Val(Mid$(FESOL$, 4, 2))
             CI% = CVI(Mid$(X$, 3, 2))
             If CI% > 0 And CI% <= NUMAS% Then
               TOCON(CODIPE% - DESDEMES%, CI%) = TOCON(CODIPE% - DESDEMES%, CI%) + CAPED
               ACCLI%(CI%) = 1
             End If
           End If
         End If
       End If
     End If
   Next U&
   Call CIERRARCH("MOVISTO")
   '
   REBLA$ = REGBLAN$("STATSTOK")
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
         ESTIM = Int((4 * SUTOTD - 2 * SUTOTI) / CAMESES% + 0.5)
         If ESTIM < 0 Then ESTIM = 0
         If ESTIM < ESTIVE + ESTIFA Then ESTIM = ESTIVE + ESTIFA
         T1$ = MKS$(MINITOT) + MKS$(MAXITOT) + MKS$((SUCOVEN + SUCOFAB) / CAMESES%) + MKS$(ESTIM)
         '
         Call REPLA(Y$, T1$, 65, 16)
         Call REPLA(Y$, Y0$, 129, 96)
         Call REPLA(Y$, Y1$, 225, 16)
         Call REPLA(Y$, Z0$, 257, 96)
         Call REPLA(Y$, Z1$, 353, 16)
         JJ& = JJ& + 1
         Call GRAREG("STATSTOK", Y$, JJ&)
       End If
     End If
   Next U&
   '
   Call SETULTREG("STATSTOK", JJ&)
   '
   Call COCOBREC
   '
   TTXX$ = Space$(40)
   Call REPLA(TTXX$, DESMES$, 1, 16)
   Call REPLA(TTXX$, HASMES$, 17, 16)
   Call REPLA(TTXX$, TRIM$(Str$(CAMESES%)), 33, 2)
   Call GRACONTROL("ESTACONS", "PERICALC", TTXX$)
   '
   TTXX$ = Format(FileDateTime(LUDAT$ + "STATSTOK.DAT"), "dd/mm/yy hh:mm")
   FEX = FECHANUM(Left$(TTXX$, 8))
   Label3 = DSEMIN$(FEX) + " " + Left$(TTXX$, 8) + " - " + Mid$(TTXX$, 10) + " hs"
   label1 = TRIM$(DESMES$)
   Label9 = TRIM$(HASMES$)
   Label12 = TRIM$(Str$(CAMESES%))
   '
90 Screen.MousePointer = 1
   '
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

