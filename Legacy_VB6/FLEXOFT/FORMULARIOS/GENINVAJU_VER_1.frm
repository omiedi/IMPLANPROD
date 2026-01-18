VERSION 5.00
Begin VB.Form GENINVAJU_VER_1 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generacion de Inventarios a Fecha"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6735
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command4 
      Caption         =   "Fecha"
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
      Left            =   5670
      Picture         =   "GENINVAJU_VER_1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   2520
      Width           =   855
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   5670
      ScaleHeight     =   75
      ScaleWidth      =   810
      TabIndex        =   6
      Top             =   2205
      Width           =   870
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5250
      TabIndex        =   0
      Top             =   210
      Width           =   175
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Avance de Proceso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4635
      Left            =   210
      TabIndex        =   4
      Top             =   735
      Width           =   5265
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Valorización de Partidas Obsoletas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   12
         Left            =   630
         TabIndex        =   20
         Top             =   4200
         Width           =   3900
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Lectura Tabla de Obsoletos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   11
         Left            =   630
         TabIndex        =   19
         Top             =   3885
         Width           =   3900
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Valorización de Inventarios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   10
         Left            =   630
         TabIndex        =   18
         Top             =   3570
         Width           =   3900
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Asociacion Recepciones con Despachos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   9
         Left            =   630
         TabIndex        =   17
         Top             =   3255
         Width           =   3900
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Apareando Despachos de Importación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   8
         Left            =   630
         TabIndex        =   16
         Top             =   2940
         Width           =   3900
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Eliminacion de Partidas Consumidas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   630
         TabIndex        =   15
         Top             =   2625
         Width           =   3900
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Procesando Devolucion - Ajuste - Transferencia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   630
         TabIndex        =   14
         Top             =   2310
         Width           =   4530
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Ordenamiento por Código de Articulo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   630
         TabIndex        =   13
         Top             =   1995
         Width           =   3900
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Neteando Partidas por Existencia Final"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   630
         TabIndex        =   12
         Top             =   1680
         Width           =   3900
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Generación de Partidas por Ingresos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   630
         TabIndex        =   11
         Top             =   1365
         Width           =   3900
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Existencia Final a la Fecha de Cálculo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   630
         TabIndex        =   10
         Top             =   1050
         Width           =   3900
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Determinación del Origen de Cada Pieza"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   630
         TabIndex        =   9
         Top             =   735
         Width           =   3900
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Tabla Cotizacion Diaria Dolar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   630
         TabIndex        =   8
         Top             =   420
         Width           =   3900
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Start"
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
      Left            =   5670
      Picture         =   "GENINVAJU_VER_1.frx":1632
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1155
      Width           =   855
   End
   Begin VB.CommandButton command1 
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
      Left            =   5670
      Picture         =   "GENINVAJU_VER_1.frx":193C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   210
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   5145
      Picture         =   "GENINVAJU_VER_1.frx":1A86
      Top             =   4830
      Width           =   2010
   End
   Begin VB.Label Label1 
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
      Height          =   330
      Left            =   4200
      TabIndex        =   5
      Top             =   210
      Width           =   1065
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de Inventario:"
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
      Left            =   2205
      TabIndex        =   3
      Top             =   315
      Width           =   1905
   End
End
Attribute VB_Name = "GENINVAJU_VER_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Hide
End Sub
'
Private Sub Command2_Click()
   Command2.Enabled = False
   If FECHANUM(Label1) > 1 Then
     Call GENINVE
   End If
   Command2.Enabled = True
End Sub
'
Private Sub Command3_Click()
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   Call SELECHO("SELFECHA")
   FECHINVE$ = VALACT1$("SELFECHA")
   If FECHANUM(FECHINVE$) > 0 Then
     Label1 = FECHINVE$
     Call GRACONTROL("INVESABO", "FECHINVE", FECHINVE$)
   End If
   Command4.Enabled = True
End Sub

'
Private Sub Form_Load()
   FECHINVE$ = CONTROL$("INVESABO", "FECHINVE")
   Label1 = FECHINVE$
End Sub

Sub GENINVE()
    '
    Dim BULEIN As String * 132
    Dim SALPIEZAS(32767)
    Dim TIENESTOCK%(32767)
    Dim CAOBSOLE(32767)
    Dim ORIGEN%(32767)
    Dim COTIDOL(32767)
    Dim COTIDO(12)
    Dim RECOSNA&(32767)
    Dim INDAJ#(256, 12)
    Dim PORNA(256), IAJUX%(256)
    '
    FECHINVE$ = Label1
    FECALCU% = FECHANUM(FECHINVE$)
    If FECALCU% < 1 Then
      Exit Sub
    End If
    FE311202% = FECALCU%     ' FECHANUM("31/12/02")
    '
' tabla cotizaciones dolar
    '
    TIVAIMP% = 1
    ARESAN$ = "ACOPROP"
    FORINVAJU.Label3 = "Históricos"
    'TIVAIMP% = COMALTER%("Tipo de Valores a Utilizar:\\Históricos\Pre-Ajustados")
    'If TIVAIMP% < 1 Or TIVAIMP% > 2 Then Exit Sub
    '
    'If TIVAIMP% = 1 Then
    '    ARESAN$ = "COSPROPR"
    '    FORINVAJU.Label3 = "Históricos"
    '  Else
    '    ARESAN$ = "COSPRAJU"
    '    FORINVAJU.Label3 = "Ajustados"
    'End If
    FORINVAJU.Label3.Refresh
    Call GRACONTROL("INVENAJU", "TIVAPROP", FORINVAJU.Label3)
    '
    FECHFIN% = FECALCU%       ' FECHANUM("31/12/02")
    INIANO$ = "01/01" + Mid$(FECHINVE$, 6)
    FECHINI% = DIANTE(FECHANUM(INIANO$))
    '
    COTIDO(0) = 1
    For U& = 1 To ULTREG&("CODIADOL")
      XX$ = REGLEIDO$("CODIADOL", U&)
      FECOT = CVI(Left$(XX$, 2))
      COTIDOX = CVD(Mid$(XX$, 3, 8))
      If COTIDOX > 0.005 Then
        If FECOT > FECHINI% Then
          If FECOT <= FECHFIN% Then
            MECOT = Val(Mid$(FECHATEX$(FECOT), 4, 2))
            If MECOT > 0 Then
              If MECOT <= 12 Then
                COTIDO(MECOT) = COTIDOX
              End If
            End If
          End If
        End If
      End If
    Next U&
    '
    ' FECHINI% = FECHANUM("31/12/01")
    FECHAN% = FECHINI%
    XX$ = REGLEIDO$("CODIADOL", 1)
      FECOTI% = CVI(Left$(XX$, 2))
      COTIDOX = CVD(Mid$(XX$, 3, 8))
    For U& = 1 To 32767: COTIDOL(U&) = COTIDOX: Next U&
    For U& = 1 To ULTREG&("CODIADOL")
      XX$ = REGLEIDO$("CODIADOL", U&)
      FECOTI% = CVI(Left$(XX$, 2))
      If FECOTI% > FECHINI% Then
        COTIDOX = CVD(Mid$(XX$, 3, 8))
        For KKI% = FECOTI% To 32766
          COTIDOL(KKI%) = COTIDOX
        Next KKI%
        FECHAN% = FECOTI%
      End If
    Next U&
    '
    EXTE$ = "DAT"   ' : If EJER% < 0 Then EXTE$ = "OLD"
    N1% = FILEOPEN%(LUDAT$ + "INDAJU." + EXTE$, 0, 132)
    CAREIN& = LOF(N1%) / 132
    For U& = 1 To CAREIN&
      Get #N1%, U&, BULEIN$
      TTXX$ = BULEIN$
      INDA1# = 0
      For KKI% = 0 To 12
        INDA2# = CVD(Mid$(BULEIN$, 21 + 8 * KKI%, 8))
        If INDA2# > 0.0001 Then INDA1# = INDA2#
        INDAJ#(U&, KKI%) = INDA1#
      Next KKI%
    Next U&
    Close #N1%
    '
    PORNA(0) = 0
    For U& = 1 To ULTREG&("GRUMAIMP")
      XX$ = REGLEIDO$("GRUMAIMP", U&)
      If U& <= 256 Then
        PORNA(U&) = CVS(Mid$(XX$, 35, 4))
        IAJUX%(U&) = Asc(Mid$(XX$, 33, 1))
      End If
    Next U&
    Call CIERRARCH("GRUMAIMP")
    YACAIN% = 1
    '
    Check1(0).Value = 1
    
    HOII% = HOY(HOS$)
    Call GRACONTROL("INVENAJU", "FECHAINV", FECHINVE$)
    Call GRACONTROL("INVENAJU", "FEGENINV", "Generado el " + HOS$ + " - " + Left$(Time$, 5) + " hs.")
    FF% = FECHANUM(FECHINVE$)
    '
    Screen.MousePointer = 11
    Call COPYSTRU("COSPROPR", "COSMEIMP")
    '
100 Caption = "Generacion de Inventarios al " + FECHINVE$
    '
'MsgBox "Voy a 500"
'GoTo 500
    ' ******* determinando origen de cada pieza ****************
    '
    FIN& = NUMAS%
    For U& = 1 To NUMAS%
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("MASTER", U& + 1)
      CI% = CVI(Left$(XX$, 2))
      CLASI$ = Mid$(XX$, 69, 8)
      ORIGEN%(CI%) = Asc(Mid$(CLASI$, 4, 1))
    Next U&
    Check1(1).Value = 1
    '
    ' ******** re-calculando stock a la fecha de calculo *****
    '
    FIN& = ULTREG&("MOVISTO")
    For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      MVS$ = REGLEIDO$("MOVISTO", I&)
      If CVI(Left$(MVS$, 2)) <= FF% Then
        CI% = CVI(Mid$(MVS$, 3, 2))
        CANTI = CVD(Mid$(MVS$, 88, 8)) - CVD(Mid$(MVS$, 96, 8))
        '
        If CI% > 0 Then
          If CI% <= NUMAS% Then
            SALPIEZAS(CI%) = SALPIEZAS(CI%) + CANTI
          End If
        End If
      End If
    Next I&
    Check1(2).Value = 1
    '
    ' ********* determinando qué piezas tienen stock **********

    KKJ% = 0
    For CI% = 1 To NUMAS%
      If SALPIEZAS(CI%) >= 0.05 Then
        TIENESTOCK%(CI%) = 1
        CDXT$ = CODEXT$(CI%)
        If Mid$(CDXT$, 16, 1) = Chr$(96) Then
           TIENESTOCK%(CI%) = 0 ' suprimir piezas de baja
        End If
        KKJ% = KKJ% + 1
      End If
    Next CI%
    '
    ' ******** generando partidas ingresadas **************
    '
    FIN& = ULTREG&("MOVISTO")
    REBLA$ = REGBLAN$("PARTIDAS")
    JEQUI& = 0
    FECAL1% = FECHANUM("11/03/02")
    FECAL2% = FECHANUM("08/03/02")
    '
    For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      MVS$ = REGLEIDO$("MOVISTO", I&)
      FEMOVI% = CVI(Left$(MVS$, 2))
      CODEQUI% = 0
      If FEMOVI% <= FF% Then
        CI% = CVI(Mid$(MVS$, 3, 2))
        If TIENESTOCK%(CI%) > 0 Then
          CAENTRA = CVD(Mid$(MVS$, 88, 8))
          If CAENTRA >= 0.05 Then
            IDENMO$ = Mid$(MVS$, 23, 10)
            If Left$(IDENMO$, 2) = "TR" Then
              XANTE$ = REGLEIDO$("MOVISTO", I& - 1)
              CIANTE% = CVI(Mid$(XANTE$, 3, 2))
              If CIANTE% <> CI% Then
                CANTE = CVD(Mid$(XANTE$, 96, 8))
                If Abs(CAENTRA - CANTE) < 0.05 Then
                  CODEQUI% = CIANTE%
                End If
              End If
            End If
            '
            ' ignorar movimientos de entrada de pr.produccion
            ' del 11/03/02 para piezas 'Tnnn' y del 8/03/02
            ' para piezas 'Wnnn'
            '
            If ORIGEN%(CI%) = 1 Then
              If FEMOVI% = FECAL1% Then
                  CDXT$ = CODEXT$(CI%)
                  If Left$(UCase$(CDXT$), 1) = "T" Then GoTo 119
                ElseIf FEMOVI% = FECAL2% Then
                  CDXT$ = CODEXT$(CI%)
                  If Left$(UCase$(CDXT$), 1) = "W" Or Left$(UCase$(CDXT$), 1) = "T" Then GoTo 119
              End If
            End If
            '
            ZZ$ = REBLA$
            Call REPLA(ZZ$, MKI$(CI%), 1, 2)
            Call REPLA(ZZ$, Chr$(ORIGEN%(CI%)), 3, 1)
            Call REPLA(ZZ$, MKI$(FEMOVI%), 5, 2)
            Call REPLA(ZZ$, IDENMO$, 7, 10)
            Call REPLA(ZZ$, MKS$(CAENTRA), 49, 4)
            Call REPLA(ZZ$, MKS$(CAENTRA), 53, 4)
            Call REPLA(ZZ$, MKI$(CODEQUI%), 63, 2)
            KKL& = KKL& + 1
            Call GRAREG("PARTIDAS", ZZ$, KKL&)
            '
          End If
        End If
      End If
119 Next I&
    '
    Call SETULTREG("PARTIDAS", KKL&)
    Call CIERRARCH("PARTIDAS")
    Check1(3).Value = 1
    '
    ' ******** neteando partidas a la fecha de calculo **************
    '
    KKL& = ULTREG&("PARTIDAS")
    For I& = KKL& To 1 Step -1
      Call TRACE(20, KKL& - I& + 1, KKL&)
      ZZ$ = REGLEIDO$("PARTIDAS", I&)
      CI% = CVI(Left$(ZZ$, 2))
      SALPARTI = CVS(Mid$(ZZ$, 53, 4))
      If SALPARTI >= SALPIEZAS(CI%) Then
          SALPARTI = SALPIEZAS(CI%)
          SALPIEZAS(CI%) = 0
          Call REPLA(ZZ$, MKS$(SALPARTI), 53, 4)
          Call GRAREG("PARTIDAS", ZZ$, I&)
        Else
          SALPIEZAS(CI%) = SALPIEZAS(CI%) - SALPARTI
      End If
    Next I&
    Call CIERRARCH("PARTIDAS")
    Check1(4).Value = 1
    '
    ' ******** ordenando partidas por codigo **************
    '
200 Call FILESORT("PARTIDAS", "PARTIDAS", 1, 5)
    Check1(5).Value = 1
    '
    ' ******** eliminando altas partidas por devolucion y ajuste **************
    '
    FIN& = ULTREG("PARTIDAS")
    JKK& = FIN&
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      ZZ$ = REGLEIDO$("PARTIDAS", U&)
      CI% = CVI(Left$(ZZ$, 2))
      TIMO$ = UCase$(Mid$(ZZ$, 7, 2))
      If TIMO$ = "DV" Or TIMO$ = "AS" Or (TIMO$ = "OF" And Asc(Mid$(ZZ$, 3, 1)) = 2) Then
        CADEVO = CVS(Mid$(ZZ$, 53, 4))
        If CADEVO >= 0.05 Then
          PRIPAR& = 0
          For JJ& = U& - 1 To 1 Step -1
            YY$ = REGLEIDO$("PARTIDAS", JJ&)
            If CVI(Left$(YY$, 2)) <> CI% Then GoTo 220
            If UCase$(Mid$(YY$, 7, 2)) <> "DV" Then
              If UCase$(Mid$(YY$, 7, 2)) <> "AS" Then
                If UCase$(Mid$(YY$, 7, 2)) <> "OF" Or Asc(Mid$(YY$, 3, 1)) <> 2 Then
                  PRIPAR& = JJ&
                  CONSUPAR = CVS(Mid$(YY$, 49, 4)) - CVS(Mid$(YY$, 53, 4))
                  If CONSUPAR >= CADEVO Then
                      SALPAR = CVS(Mid$(YY$, 53, 4)) + CADEVO
                      Call REPLA(YY$, MKS$(SALPAR), 53, 4)
                      Call GRAREG("PARTIDAS", YY$, JJ&)
                      ZZ$ = REGLEIDO$("PARTIDAS", U&)
                      Call REPLA(ZZ$, MKS$(0), 53, 4)
                      Call GRAREG("PARTIDAS", ZZ$, U&)
                      CADEVO = 0
                      GoTo 290
                    Else
                      CADEVO = CADEVO - CONSUPAR
                      SALPAR = CVS(Mid$(YY$, 49, 4))
                      Call REPLA(YY$, MKS$(SALPAR), 53, 4)
                      Call GRAREG("PARTIDAS", YY$, JJ&)
                  End If
                End If
              End If
            End If
          Next JJ&
220       If CADEVO >= 0.05 Then
            If PRIPAR& > 0 Then
              JJ& = PRIPAR&
              YY$ = REGLEIDO$("PARTIDAS", JJ&)
              SALPAR = CVS(Mid$(YY$, 53, 4)) + CADEVO
              Call REPLA(YY$, MKS$(SALPAR), 53, 4)
              Call REPLA(YY$, MKS$(SALPAR), 49, 4)
              Call GRAREG("PARTIDAS", YY$, JJ&)
              ZZ$ = REGLEIDO$("PARTIDAS", U&)
              Call REPLA(ZZ$, MKS$(0), 53, 4)
              Call GRAREG("PARTIDAS", ZZ$, U&)
              GoTo 290
            End If
            ZZ$ = REGLEIDO$("PARTIDAS", U&)
            Call REPLA(ZZ$, MKS$(CADEVO), 53, 4)
            Call GRAREG("PARTIDAS", ZZ$, U&)
          End If
          GoTo 290
        End If
      End If
      '
      If TIMO$ = "AJ" Then
        If CVI(Mid$(ZZ$, 5, 2)) <= FECHANUM("31/07/98") Then
          CADEVO = CVS(Mid$(ZZ$, 53, 4))
          If CADEVO >= 0.05 Then
            RFF$ = RECFILT$("DESPAIMP", 1, MKI$(CI%))
            For KU& = Len(RFF$) - 3 To 1 Step (-4)
              REDES& = CVS(Mid$(RFF$, KU&, 4))
              XX$ = REGLEIDO$("DESPAIMP", REDES&)
              FEI% = CVI(Mid$(XX$, 3, 2))
              If FEI% > 0 Then
                If FEI% <= FECHANUM("31/07/02") Then
                  CANI = CVD(Mid$(XX$, 33, 8))
                  If CANI > 0 Then
                    DPSX$ = TRIM$(Mid$(XX$, 17, 16)) + " - " + TRIM$(Mid$(XX$, 5, 12))
                    Call REPLA(ZZ$, DPSX$, 17, 32)
                    Call GRAREG("PARTIDAS", ZZ$, U&)
                    GoTo 290   ' LO ENCONTRO
                  End If
                End If
              End If
            Next KU&
          End If
        End If
      End If
      '
      If TIMO$ = "TR" Then
        If CVS(Mid$(ZZ$, 53, 4)) >= 0.05 Then
          CODEQUI% = CVI(Mid$(ZZ$, 63, 2))
          If CODEQUI% > 0 Then
            RFF$ = RECFILT$("PARTIDAS", 1, MKI$(CODEQUI%))
            For KQL& = Len(RFF$) - 3 To 1 Step -4
              U1& = CVS(Mid$(RFF$, KQL&, 4))
              TTT$ = REGLEIDO$("PARTIDAS", U1&)
              If CVI(Left$(TTT$, 2)) = CODEQUI% Then
                If CVS(Mid$(TTT$, 49, 4)) > CVS(Mid$(TTT$, 53, 4)) + 0.1 Then
                  FEPARTI% = CVI(Mid$(TTT$, 5, 2))
                  Call REPLA(ZZ$, MKI$(FEPARTI%), 5, 2)
                  Call GRAREG("PARTIDAS", ZZ$, U&)
                  GoTo 290
                End If
              End If
            Next KQL&
          End If
        End If
      End If
290 Next U&
    Call CIERRARCH("PARTIDAS")
    Check1(6).Value = 1
    '
    ' ******** eliminando partidas consumidas **************
    '
    Call COPYSTRU("PARTIDAS", "PARTIDES")
    JJ& = 0
    KKL& = ULTREG&("PARTIDAS")
    For I& = 1 To KKL&
      Call TRACE(20, I&, KKL&)
      ZZ$ = REGLEIDO$("PARTIDAS", I&)
      If CVS(Mid$(ZZ$, 53, 4)) >= 0.05 Then
        JJ& = JJ& + 1
        Call GRAREG("PARTIDES", ZZ$, JJ&)
      End If
    Next I&
    Call SETULTREG("PARTIDES", JJ&)
    Call CIERRARCH("PARTIDES")
    '
    ' ******** ordenando por fecha **************
    '
    Call FILESORT("PARTIDES", "PARTIDIS", 1, 5)
    Check1(7).Value = 1
    '
    ' ******** Despachos de importacion **************
    '
300 FIN& = ULTREG("PARTIDIS")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      ZZ$ = REGLEIDO$("PARTIDIS", U&)
      CI% = CVI(Left$(ZZ$, 2))
      TIMO$ = UCase$(Mid$(ZZ$, 7, 2))
      CAORI = CVS(Mid$(ZZ$, 49, 4))
      If TIMO$ = "BR" Then
        If TRIM$(Mid$(ZZ$, 17, 32)) = "" Then
          FEMORI% = CVI(Mid$(ZZ$, 5, 2))
          '
          RFF$ = RECFILT$("DESPAIMP", 1, MKI$(CI%))
          For KU& = 1 To Len(RFF$) Step 4
            REDES& = CVS(Mid$(RFF$, KU&, 4))
            XX$ = REGLEIDO$("DESPAIMP", REDES&)
            FEI% = CVI(Mid$(XX$, 3, 2))
            If FEI% > 0 Then
              If Abs(FEMORI% - FEI%) < 15 Then
                CANI = CVD(Mid$(XX$, 33, 8))
                If Abs(CAORI - CANI) < 1 Then
                  DPSX$ = TRIM$(Mid$(XX$, 17, 16)) + " - " + TRIM$(Mid$(XX$, 5, 12))
                  Call REPLA(ZZ$, DPSX$, 17, 32)
                  Call GRAREG("PARTIDIS", ZZ$, U&)
                  GoTo 390   ' LO ENCONTRO
                End If
              End If
            End If
          Next KU&
        End If
      End If
      '
      If TIMO$ = "TR" Then
        If TRIM$(Mid$(ZZ$, 17, 32)) = "" Then
          FEMORI% = CVI(Mid$(ZZ$, 5, 2))
          If CVI(Mid$(ZZ$, 63, 2)) > 0 Then
            CI% = CVI(Mid$(ZZ$, 63, 2))
          End If
          '
          RFF$ = RECFILT$("DESPAIMP", 1, MKI$(CI%))
          For KU& = 1 To Len(RFF$) Step 4
            REDES& = CVS(Mid$(RFF$, KU&, 4))
            XX$ = REGLEIDO$("DESPAIMP", REDES&)
            FEI% = CVI(Mid$(XX$, 3, 2))
            If FEI% > 0 Then
              If Abs(FEMORI% - FEI%) < 15 Then
                CANI = CVD(Mid$(XX$, 33, 8))
                If CANI > 0.05 Then
                  DPSX$ = TRIM$(Mid$(XX$, 17, 16)) + " - " + TRIM$(Mid$(XX$, 5, 12))
                  Call REPLA(ZZ$, DPSX$, 17, 32)
                  Call GRAREG("PARTIDIS", ZZ$, U&)
                  GoTo 390   ' LO ENCONTRO
                End If
              End If
            End If
          Next KU&
        End If
      End If
      
390 Next U&
    Check1(8).Value = 1
    '
    ' ********* buscando por numero de BR **************
    '
400 FIN& = ULTREG("PARTIDIS")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      ZZ$ = REGLEIDO$("PARTIDIS", U&)
      TIMO$ = UCase$(Mid$(ZZ$, 7, 2))
      If TIMO$ = "BR" Then
        If TRIM$(Mid$(ZZ$, 17, 32)) = "" Then
          IDEMO$ = TRIM$(Mid$(ZZ$, 7, 10))
          If IDEMO$ <> "" Then
            For JJ& = 1 To FIN&
              YY$ = REGLEIDO$("PARTIDIS", JJ&)
              If TRIM$(Mid$(YY$, 7, 10)) = IDEMO$ Then
                DSPI$ = Mid$(YY$, 17, 32)
                If TRIM$(DSPI$) <> "" Then
                  Call REPLA(ZZ$, DSPI$, 17, 32)
                  Call GRAREG("PARTIDIS", ZZ$, U&)
                  GoTo 490
                End If
              End If
            Next JJ&
          End If
        End If
      End If
490 Next U&
    Call CIERRARCH("PARTIDIS")
    Check1(9).Value = 1
          '
500 FIN& = ULTREG&("PARTIDIS")
    IMES% = Val(Mid$(FECHINVE$, 4, 2))
    If FECALCU% > FE311202% Then IMES% = 12
    ANOCA$ = Mid$(FECHINVE$, 7, 2)
    KMI& = 0
    '
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("PARTIDIS", U&)
      If Asc(Mid$(XX$, 3, 1)) = 2 Then ' SOLO IMPORTADOS
        CI% = CVI(Left$(XX$, 2))
        COEQUI% = CVI(Mid$(XX$, 63, 2))
        SALPARTI = CVS(Mid$(XX$, 53, 4))
        FEALTA% = CVI(Mid$(XX$, 5, 2))
        VAFOB# = 0: GRUMA% = 0
        '
        GoTo 502      ' NUEVO CRITERIO PARA 2003 SEGUN MINGOTE
        ' valorizar en dolares para inventario al 30/06/03
        ' si partida anterior al año de calculo,
        ' tomar valor de inventario anterior sin gastos,
        ' no el de b.datos productos importados
        FEALTAX = FEALTA%
        FECHAX$ = FECHATEX$(FEALTAX)
        If Mid$(FECHAX$, 7, 2) <> ANOCA$ Then
          'AAA = bbb
          GoTo 505
        End If
        '
502     RFF$ = RECFILT$("FOBMAIMP", 1, MKI$(CI%))
        For KKI% = 1 To Len(RFF$) Step 4
          REFO& = CVS(Mid$(RFF$, KKI%, 4))
          YY$ = REGLEIDO$("FOBMAIMP", REFO&)
          GRUMA% = CVI(Mid$(YY$, 3, 2)) Mod 256
          VAFOB# = CVD(Mid$(YY$, 5, 8))
          If VAFOB# > 0.00005 Then GoTo 510
        Next KKI%
        '
        If COEQUI% > 0 Then
          If COEQUI% <> CI% Then
            RFF$ = RECFILT$("FOBMAIMP", 1, MKI$(COEQUI%))
            For KKI% = 1 To Len(RFF$) Step 4
              REFO& = CVS(Mid$(RFF$, KKI%, 4))
              YY$ = REGLEIDO$("FOBMAIMP", REFO&)
              GRUMA% = CVI(Mid$(YY$, 3, 2)) Mod 256
              VAFOB# = CVD(Mid$(YY$, 5, 8))
              If VAFOB# > 0.00005 Then GoTo 510
            Next KKI%
          End If
        End If
        '
        'tomar valores del inventario anterior dolar 1 a 1
505 'abc$ = CODEXT$(CI%)
        RFF$ = RECFILT$("VUNINVAN", 1, MKI$(CI%))
        If Len(RFF$) >= 4 Then
          REFO& = CVS(Left$(RFF$, 4))
          YY$ = REGLEIDO$("VUNINVAN", REFO&)
          VAFOB# = CVD(Mid$(YY$, 3, 8))
          GRUMA% = 0
          GoTo 510
        End If
        '
'MsgBox "Falta Valor FOB de " + TRIM$(CODEXT$(CI%))
        '
510     FEALTAX = CVI(Mid$(XX$, 5, 2))
        FECHAX$ = FECHATEX$(FEALTAX)
        If FEALTAX > FE311202% Then
            IMORI% = 12
          ElseIf Mid$(FECHAX$, 7, 2) <> ANOCA$ Then
            IMORI% = 0
          Else
            IMORI% = Val(Mid$(FECHAX$, 4, 2))
        End If
        PORNACIO = PORNA(GRUMA%)
        FEALTAI% = FEALTA%: If FEALTAI% < FECHINI% Then FEALTAI% = FECHINI%
        CTDL = COTIDOL(FEALTAI%)    '  If FEALTA% > FE311202% Then CTDL = 2.8
        Call REPLA(XX$, MKD$(VAFOB#), 65, 8)
        Call REPLA(XX$, MKS$(CTDL), 73, 4)
        Call REPLA(XX$, MKS$(PORNACIO), 77, 4)
        COEFI = CTDL * (1 + PORNACIO / 100)
        Call REPLA(XX$, MKD$(VAFOB# * COEFI), 81, 8)
        Call REPLA(XX$, MKD$(VAFOB# * COEFI * SALPARTI), 89, 8)
        '
        VUACT# = VAFOB# * COEFI * INDAJ#(1, IMES%) / INDAJ#(1, IMORI%)
'If FECALCU% > FE311202% Then VUACT# = VAFOB# * COEFI * 2.8 / INDAJ(1, IMORI%)
        Call REPLA(XX$, MKD$(VUACT#), 153, 8)
        Call REPLA(XX$, MKD$(VUACT# * SALPARTI), 161, 8)
        '
        CI% = CVI(Left$(XX$, 2))
        VANERE# = 0
        VANERE# = PRELISTA#("LIPRE003", CI%) * (1 - 0.32)
        If VANERE# < 0.005 Then
          VANERE# = PRELISTA#("LIPRE001", CI%) * (1 - 0.32)
        End If
        '
        Call GRAREG("PARTIDIS", XX$, U&)
        '
        CI% = CVI(Left$(XX$, 2))
        CTDM = COTIDOL(FF%)      '  If FF% > FE311202% Then CTDM = 2.8
        COEFJ = CTDM * (1 + PORNACIO / 100)
        For JJ& = U& To 1 Step -1
          YY$ = REGLEIDO$("PARTIDIS", JJ&)
          If CVI(Left$(YY$, 2)) <> CI% Then GoTo 520
          SALPARTJ = CVS(Mid$(YY$, 53, 4))
          Call REPLA(YY$, MKD$(VAFOB# * COEFI), 105, 8)
          Call REPLA(YY$, MKD$(VAFOB# * COEFI * SALPARTJ), 113, 8)
          '
          Call REPLA(YY$, MKD$(VAFOB# * COEFJ), 129, 8)
          Call REPLA(YY$, MKD$(VAFOB# * COEFJ * SALPARTJ), 137, 8)
          '
          If VANERE# < 0.005 Then VANERE# = VAFOB# * COEFJ
          Call REPLA(YY$, MKD$(VANERE#), 177, 8)
          Call REPLA(YY$, MKD$(VANERE# * SALPARTJ), 185, 8)
          '
          Call GRAREG("PARTIDIS", YY$, JJ&)
          '
        Next JJ&
        '
520     RFG$ = RECFILT$("COSMEIMP", 1, MKI$(CI%))
        REMEI& = 0
        If Len(RFG$) >= 4 Then
          REMEI& = CVS(Left$(RFG$, 4))
        End If
        If REMEI& < 1 Or REMEI& > ULTREG&("COSMEIMP") Then
          KMI& = KMI& + 1
          REMEI& = KMI&
          ZY$ = REGBLAN$("COSMEIMP")
          Call REPLA(ZY$, MKI$(CI%), 1, 2)
          Call GRAREG("COSMEIMP", ZY$, REMEI&)
        End If
        '
        ZY$ = REGLEIDO$("COSMEIMP", REMEI&)
        '
        For KKUI% = IMORI% To 12
          VAMEN# = VAFOB# * (1 + PORNACIO / 100) * COTIDO(KKUI%)
          Call REPLA(ZY$, MKD$(VAMEN#), 9 + 8 * KKUI%, 8)
        Next KKUI%
        '
        For KKUI% = 0 To IMORI% - 1
          If CVD(Mid$(ZY$, 9 + 8 * KKUI%, 8)) < 0.00005 Then
            VAMEN# = VAFOB# * (1 + PORNACIO / 100) * COTIDO(KKUI%)
            Call REPLA(ZY$, MKD$(VAMEN#), 9 + 8 * KKUI%, 8)
          End If
        Next KKUI%
        '
        Call GRAREG("COSMEIMP", ZY$, REMEI&)
        '
        GoTo 590
      End If
      '
      ' *********** NACIONALES - PROPIA PRODUCCION
      '
      CI% = CVI(Left$(XX$, 2))
      COEQUI% = CVI(Mid$(XX$, 63, 2))
      SALPARTI = CVS(Mid$(XX$, 53, 4))
      FEALTA% = CVI(Mid$(XX$, 5, 2))
      VAHIS# = 0: GRUMA% = 0
      If RECOSNA&(CI%) < 1 Then
        RFF$ = RECFILT$(ARESAN$, 1, MKI$(CI%))
        If Len(RFF$) < 4 Then
          TCPX$ = REGBLAN$(ARESAN$)
          MONECO% = MONECOSTO%(CI%)
          COUN# = COSUNI(CI%) / TICAMONE(MONECO%)
          Call REPLA(TCPX$, MKI$(CI%), 1, 2)
          Call REPLA(TCPX$, MKD$(COUN#), 3, 8)
          Call REGAPP(ARESAN$, TCPX$)
          Call CIERRARCH(ARESAN$)
          REFO& = ULTREG&(ARESAN$)
          RFF$ = MKS$(REFO&)
        End If
        REFO& = CVS(Left$(RFF$, 4))
        RECOSNA&(CI%) = REFO&
      End If
      '
      REFO& = RECOSNA&(CI%)
      YY$ = REGLEIDO$(ARESAN$, REFO&)
      FEALTAX = CVI(Mid$(XX$, 5, 2))
      FECHAX$ = FECHATEX$(FEALTAX)
      If FECALCU% > FE311202% Then
          IMORI% = 12
        ElseIf Mid$(FECHAX$, 7, 2) <> ANOCA$ Then
          IMORI% = 0
        Else
          IMORI% = Val(Mid$(FECHAX$, 4, 2))
      End If
      VAHIS# = CVD(Mid$(YY$, 3, 8))
      VAREP# = CVD(Mid$(YY$, 3, 8))
      '
      PORNACIO = 0
      Call REPLA(XX$, MKD$(VAHIS#), 65, 8)
      Call REPLA(XX$, MKS$(1), 73, 4)
      Call REPLA(XX$, MKS$(PORNACIO), 77, 4)
      COEFI = 1
      Call REPLA(XX$, MKD$(VAHIS# * COEFI), 81, 8)
      Call REPLA(XX$, MKD$(VAHIS# * COEFI * SALPARTI), 89, 8)
      '
      VUACT# = VAHIS# * COEFI * INDAJ#(1, IMES%) / INDAJ#(1, IMORI%)
      ' If FECALCU% > FE311202% Then VUACT# = VAHIS# * COEFI
      Call REPLA(XX$, MKD$(VUACT#), 153, 8)
      Call REPLA(XX$, MKD$(VUACT# * SALPARTI), 161, 8)
      '
      CI% = CVI(Left$(XX$, 2))
      VANERE# = 0
      VANERE# = PRELISTA#("LIPRE003", CI%) * (1 - 0.32)
      If VANERE# < 0.005 Then
         VANERE# = PRELISTA#("LIPRE001", CI%) * (1 - 0.32)
      End If
      '
      Call GRAREG("PARTIDIS", XX$, U&)
      '
      CI% = CVI(Left$(XX$, 2))
      CTDM = COTIDOL(FF%)   ' : If FF% > FE311202% Then CTDM = 2.8
      COEFJ = CTDM * (1 + PORNACIO / 100)
      For JJ& = U& To 1 Step -1
        '
        YY$ = REGLEIDO$("PARTIDIS", JJ&)
        If CVI(Left$(YY$, 2)) <> CI% Then GoTo 590
        SALPARTJ = CVS(Mid$(YY$, 53, 4))
        Call REPLA(YY$, MKD$(VAHIS#), 105, 8)
        Call REPLA(YY$, MKD$(VAHIS# * COEFI * SALPARTJ), 113, 8)
        '
        Call REPLA(YY$, MKD$(VAREP#), 129, 8)
        Call REPLA(YY$, MKD$(VAREP# * SALPARTJ), 137, 8)
        '
        If VANERE# < 0.005 Then VANERE# = VAREP#
        Call REPLA(YY$, MKD$(VANERE#), 177, 8)
        Call REPLA(YY$, MKD$(VANERE# * SALPARTJ), 185, 8)
        '
        Call GRAREG("PARTIDIS", YY$, JJ&)
        '
      Next JJ&
      GoTo 590
      '
590 Next U&
    '
    Call CIERRARCH("PARTIDIS")
    Call SETULTREG("COSMEIMP", KMI&)
    Call CIERRARCH("COSMEIMP")
    Check1(10).Value = 1
    '
    Call HEADERS("PARTIDIS", "")
    Call HEADERS("PARTIDIS", "Inventario por Partidas al: " + FECHINVE$)
    Call CIERRARCH("*.*")
    '
    If FECALCU% > FECHANUM("31/12/02") Then GoTo 700    ' no netea obsoletos
    '
600 FIN& = ULTREG&("TAOBSOLE")
    For U& = 1 To FIN&
      XX$ = REGLEIDO$("TAOBSOLE", U&)
      CIKI% = CVI(Left$(XX$, 2))
      If CIKI% > 0 Then
        CAOBSOLE(CIKI%) = CAOBSOLE(CIKI%) + (CVD(Mid$(XX$, 3, 8)))
      End If
    Next U&
    Call CIERRARCH("TAOBSOLE")
    Check1(11).Value = 1
    '
    Call COPYSTRU("PARTIDIS", "PORTIDIS")
    JJ& = 0
    FIN& = ULTREG&("PARTIDIS")
    '
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("PARTIDIS", U&)
      CI% = CVI(Left$(XX$, 2))
      SALPARTI = CVS(Mid$(XX$, 53, 4))
      '
      If CAOBSOLE(CI%) >= SALPARTI Then
          CAOBSOLE(CI%) = CAOBSOLE(CI%) - SALPARTI
        Else
          SALPARTI = CAOBSOLE(CI%)
          CAOBSOLE(CI%) = 0
      End If
      '
      If SALPARTI >= 0.05 Then
        Call REPLA(XX$, MKS$(SALPARTI), 53, 4)
        For KKI% = 1 To 5
          ITOPART# = CVD(Mid$(XX$, 57 + 24 * KKI%, 8)) * SALPARTI
          POREPLA% = 65 + 24 * KKI%
          Call REPLA(XX$, MKD$(ITOPART#), POREPLA%, 8)
        Next KKI%
        JJ& = JJ& + 1
        Call GRAREG("PORTIDIS", XX$, JJ&)
      End If
      '
    Next U&
    '
    Call SETULTREG("PORTIDIS", JJ&)
    Call CIERRARCH("PORTIDIS")
    Check1(12).Value = 1
    '
700 Call HEADERS("PORTIDIS", "")
    Call HEADERS("PORTIDIS", "Inventario por Partidas al: " + FECHINVE$)
    Call CIERRARCH("*.*")
    '
    Screen.MousePointer = 1
    Call COMUNI("Calculo Completo")
    Hide
    '
End Sub
'

