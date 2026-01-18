VERSION 5.00
Begin VB.Form FORLICODOS 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Liquidación de Comisiones por Ventas"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9885
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   9885
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "Vend's"
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
      Left            =   5040
      Picture         =   "FORLICODOS.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   210
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Edit"
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
      Left            =   7875
      Picture         =   "FORLICODOS.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Asignar Vendedor por O/Venta"
      Top             =   210
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Listado"
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
      Left            =   6930
      Picture         =   "FORLICODOS.frx":0454
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   210
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Totales por Vendedor"
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
      Left            =   5040
      TabIndex        =   13
      Top             =   1260
      Width           =   4635
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2085
         Left            =   210
         TabIndex        =   14
         Top             =   420
         Width           =   4215
      End
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H80000016&
      Height          =   1275
      Left            =   210
      ScaleHeight     =   1215
      ScaleWidth      =   4575
      TabIndex        =   7
      Top             =   210
      Width           =   4635
      Begin VB.Label Label6 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2940
         TabIndex        =   20
         Top             =   735
         Width           =   855
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1575
         TabIndex        =   19
         Top             =   735
         Width           =   1170
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1575
         TabIndex        =   18
         Top             =   210
         Width           =   2220
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Período:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   225
         Index           =   10
         Left            =   -735
         TabIndex        =   12
         Top             =   315
         Width           =   2220
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Generado el:"
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
         Index           =   13
         Left            =   -735
         TabIndex        =   11
         Top             =   840
         Width           =   2220
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Hs"
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
         Index           =   12
         Left            =   3990
         TabIndex        =   10
         Top             =   840
         Width           =   1590
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   225
         Left            =   2835
         TabIndex        =   9
         Top             =   840
         Width           =   15
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "-"
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
         Left            =   2730
         TabIndex        =   8
         Top             =   735
         Width           =   225
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Avance de Proceso"
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
      Top             =   1680
      Width           =   4635
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   525
         ScaleHeight     =   75
         ScaleWidth      =   3525
         TabIndex        =   16
         Top             =   1785
         Width           =   3585
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   17
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Cálculo de Comisiones"
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
         Left            =   525
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1365
         Width           =   2745
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Vendedor Según Orden de Venta"
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
         Left            =   525
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1050
         Width           =   4005
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Número de Orden de Venta por Factura"
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
         Left            =   525
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   735
         Width           =   3690
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Selección de Facturación del Período"
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
         Left            =   525
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   420
         Width           =   3585
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cálculo"
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
      Left            =   5985
      Picture         =   "FORLICODOS.frx":0ABE
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   210
      Width           =   855
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
      Height          =   855
      Left            =   8820
      Picture         =   "FORLICODOS.frx":0DC8
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   210
      Width           =   855
   End
End
Attribute VB_Name = "FORLICODOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
    Check1.Value = 0
End Sub

Private Sub Check2_Click()
    Check2.Value = 0
End Sub

Private Sub Check3_Click()
    Check3.Value = 0
End Sub

Private Sub Check4_Click()
    Check4.Value = 0
End Sub

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Private Sub Command2_Click()
    '
    Command2.Enabled = False
    Call CALCOMIS
99  Command2.Enabled = True
    '
End Sub

Private Sub Command3_Click()
    '
    Command3.Enabled = False
    Call CONECRUN("*LISCODOS", "Listado Mensual de Comisiones Liquidadas")
    Command3.Enabled = True
    '
End Sub

Private Sub Command4_Click()
    Call TRALOCAL("LICODOSI", "")
    VTB% = VENTABU%("EDLICODO")
    If VTB% >= 0 Then
      Screen.MousePointer = 11
      For U& = 1 To ULTREG&("!LICODOSI")
        X$ = REGLEIDO$("!LICODOSI", U&)
        VENDE% = Asc(Mid$(X$, 41, 1))
        NPED& = CVS(Mid$(X$, 21, 4))
        PORCOMIP = CVS(Mid$(X$, 51, 4))
        If PORCOMIP < 0 Then PORCOMIP = 0
        If PORCOMIP > 99.9 Then PORCOMIP = 99.9
        Y$ = REGLEIDO$("LICODOSI", U&)
        If Asc(Mid$(Y$, 41, 1)) <> VENDE% Or Abs(CVS(Mid$(Y$, 51, 4)) - PORCOMIP) > 0.05 Then
          '
          For I& = 1 To ULTREG&("PEDDETA")
            Z$ = REGLEIDO$("PEDDETA", I&)
            NOPE& = CVS(Left$(Z$, 4))
            If NOPE& = NPED& Then
              Call REPLA(Z$, Chr$(VENDE%), 30, 1)
              Call REPLA(Z$, MKI$(10 * PORCOMIP), 31, 2)
              Call GRAREG("PEDDETA", Z$, I&)
            End If
          Next I&
          '
          For I& = 1 To ULTREG&("PEDENCA")
            Z$ = REGLEIDO$("PEDENCA", I&)
            NOPE& = CVS(Left$(Z$, 4))
            If NOPE& = NPED& Then
              Call REPLA(Z$, Chr$(VENDE%), 270, 1)
              Call GRAREG("PEDENCA", Z$, I&)
            End If
          Next I&
          '
          Call GRAREG("LICODOSI", X$, U&)
          '
        End If
      Next U&
      Call CIERRARCH("LICODOSI")
      Call CIERRARCH("!LICODOSI")
      Call CALCOMIS
      Screen.MousePointer = 1
    End If
    '
    Call CIERRARCH("LICODOSI")
    Call CIERRARCH("!LICODOSI")
    '
End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
   Call FINAL("*VENDEDOR")
   Command5.Enabled = True
End Sub

Private Sub Form_Activate()
    If List1.ListCount < 1 Then
        Command3.Enabled = False
        Command4.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    '
    Call VERJOBID
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    Show
    Refresh
    'FORSELMEC.Show 1
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CIERRARCH("*.*")
End Sub

Private Sub MESAN_GOTFOCUS()
    Command2.SetFocus
End Sub

Private Sub Option1_Click()
    If Option2.Value = False Then
        Option1.Value = False
    End If
End Sub

Private Sub Option2_Click()
    List1.Clear
    Command3.Enabled = False
    Command4.Enabled = False
End Sub

Private Sub Option3_Click()
    List1.Clear
    Command3.Enabled = False
    Command4.Enabled = False
    Option1.Value = False
End Sub

Private Sub Option4_Click()
    List1.Clear
    Command3.Enabled = False
    Command4.Enabled = False
    Option1.Value = False
End Sub

Private Sub Option5_Click()
    List1.Clear
    Command3.Enabled = False
    Command4.Enabled = False
    Option1.Value = False
End Sub

Private Sub Text4_GOTFOCUS(Index As Integer)
    Command2.SetFocus
End Sub

Private Sub List1_DblClick()
   Screen.MousePointer = 11
   CORRESP$ = AJUSTI$(FORLICODOS.Label3.Caption, 12)
   Call COPYSTRU("LICODOSI", "LICODOVE")
   VENDE% = Val(Left$(List1.TEXT, 4))
   VENDX$ = Mid$(List1.TEXT, 6, 17)
   '
   Call HEADERS("LICODOVE", "")
   Call HEADERS("LICODOVE", "Corresponde a: " + TRIM$(VENDX$) + " - " + CORRESP$)
   '
   JJ& = 0
   FIN& = ULTREG&("LICODOSI")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("LICODOSI", U&)
      If Asc(Mid$(X$, 41, 1)) = VENDE% Then
         JJ& = JJ& + 1
         Call GRAREG("LICODOVE", X$, JJ&)
      End If
   Next U&
   Call SETULTREG("LICODOVE", JJ&)
   Call CIERRARCH("LICODOVE")
   Call CIERRARCH("LICODOSI")
   '
   Call FINAL("*LICODOVE")
   Screen.MousePointer = 1
End Sub

Sub CALCOMIS()
    '
    Screen.MousePointer = 11
    '
    YALIQUI% = 0
    Dim CVX%(128), COVEN(128)
    Dim PORCO(256), TOCOMI#(256), TOCOB#(256)
    Dim REGICLI&(32767), VENDECLI%(32767)
    '
    Dim SUPED#(256), COMPED#(256)
    Dim SUCOB#(256), COMCOB#(256)
    '
    Dim NFAC&(2048), NCLI%(2048), NPED&(2048), NREM&(2048)
    Dim NVEN%(2048), INET#(2048), NOCO$(2048), FEFA%(2048)
    Dim COMIP(2048)
    
    NUVE% = 0
    YAGERE% = 0
    For I& = 1 To ULTREG&("VENDEDOR")
        X$ = REGLEIDO$("VENDEDOR", I&)
        CV% = Asc(Left$(X$, 1))
        If CV% > 0 Then
            NUVE% = NUVE% + 1
            CVX%(NUVE%) = CV%
            COVEN(NUVE%) = CVS(Mid$(X$, 44, 4))
            PORCO(CV%) = COVEN(NUVE%)
            If CV% = 99 Then YAGERE% = 1
        End If
    Next I&
    '
    If YAGERE% = 0 Then
        X$ = REGBLAN$("VENDEDOR")
        Call REPLA(X$, Chr$(99), 1, 1)
        Call REPLA(X$, "GERENCIA", 2, 30)
        Call REPLA(X$, "GERENCIA", 32, 12)
        Call REPLA(X$, MKS$(0), 44, 4)
        Call REPLA(X$, MKS$(0), 48, 4)
        Call REGAPP("VENDEDOR", X$)
        '
        NUVE% = NUVE% + 1
        CVX%(NUVE%) = 99
        COVEN(NUVE%) = CVS(Mid$(X$, 44, 4))
    End If
    Call CIERRARCH("VENDEDOR")
    '
    CORRESP$ = TRIM$(Label3)
    CORRESQ$ = MESANO$(CORRESP$, DES%, HAS%)
    FEX = DES%: DESFE$ = FECHATEX$(FEX)
    FEX = HAS%: HASFE$ = FECHATEX$(FEX)
    '
    CORRESQ$ = TRIM$(CORRESP$)
    ABREMES$ = Left$(CORRESQ$, 3) + Mid$(DESFE$, 7, 2)
    '
    Call HEADERS("LICOMIX", "")
    Call HEADERS("LICOMIX", "Corresponde a: " + CORRESP$)
    '
    J& = 0: JLIBRO& = 0
    '
    Call HEADERS("LICOMIX", "Por: Facturacion")
    '
    Check1.Value = 2
    DoEvents
    FIN& = ULTREG&("FACTUR")
    For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      X$ = REGLEIDO$("FACTUR", I&)
      FF% = CVI(Mid$(X$, 7, 2))
      If FF% >= DES% Then
        If FF% <= HAS% Then
          NRO& = CVS(Mid$(X$, 3, 4))
          If NRO& > 0 Then
            TIDOC% = Asc(Mid$(X$, 13, 1))
            If TIDOC% = 1 Or TIDOC% = 2 Or TIDOC% = 6 Then ' solo facturas y creditos
              NC% = CVI(Mid$(X$, 1, 2))
              If NC% > 0 Then
                VENDE% = VENDCLI%(NC%)
                If VENDE% < 1 Then VENDE% = 99
                ITOT# = CVD(Mid$(X$, 15, 8)) + CVD(Mid$(X$, 49, 8))
                If TIDOC% = 6 Then ITOT# = (-1) * Abs(ITOT#)
                '
                J& = J& + 1
                NFAC&(J&) = NRO&
                NCLI%(J&) = NC%
                NPED&(J&) = 0
                NREM&(J&) = 0
                NVEN%(J&) = VENDE%
                INET#(J&) = ITOT#
                FEFA%(J&) = FF%
                '
              End If
            End If
          End If
        End If
      End If
    Next I&
    '
    CAFASEL& = J&
    Check1.Value = 1
    '
    Check2.Value = 2
    DoEvents
112 FIN& = ULTREG&("MOVISTO")
    For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      XS$ = REGLEIDO$("MOVISTO", I&)
      FF% = CVI(Left$(XS$, 2))
      If FF% >= DES% Then
        If FF% <= HAS% Then
          COMO$ = UCase$(Mid$(XS$, 33, 2))
          If COMO$ = "FC" Or COMO$ = "NC" Then
            TP1% = 2: If COMO$ = "NC" Then TP1% = 6
            CLI1% = CVI(Mid$(XS$, 84, 2))
            NRO1& = Val(Mid$(XS$, 36, 7))
            NRO2& = Val(Mid$(XS$, 26, 7)) ' nro remito
            GoSub 500
          End If
        End If
      End If
    Next I&
    '
113 FIN& = ULTREG&("MOVIREPA")
    For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      XS$ = REGLEIDO$("MOVIREPA", I&)
      FF% = CVI(Left$(XS$, 2))
      If FF% >= DES% Then
        If FF% <= HAS% Then
          COMO$ = UCase$(Mid$(XS$, 33, 2))
          If COMO$ = "FC" Or COMO$ = "NC" Then
            TP1% = 2: If COMO$ = "NC" Then TP1% = 6
            CLI1% = CVI(Mid$(XS$, 84, 2))
            NRO1& = Val(Mid$(XS$, 36, 7))
            NRO2& = Val(Mid$(XS$, 26, 7)) ' nro remito
            GoSub 500
          End If
        End If
      End If
    Next I&
    Check2.Value = 1
    '
    Check3.Value = 2
    DoEvents
    FIN& = ULTREG&("PEDDETA")
    For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      X$ = REGLEIDO$("PEDDETA", I&)
      NOPE& = CVS(Left$(X$, 4))
      CLIPE% = CVI(Mid$(X$, 7, 2))
      OCOPE$ = Mid$(X$, 9, 16)
      VENPE% = Asc(Mid$(X$, 30, 1))
      RMPE& = CVS(Mid$(X$, 105, 4))
      PORCOMIPE = CVI(Mid$(X$, 31, 2)) / 10
      If PORCOMIPE < 0 Then PORCOMIPE = 0
      If PORCOMIPE > 99.9 Then PORCOMIPE = 99.9
      GoSub 550
    Next I&
    Check3.Value = 1
    '
200 J& = 0
    Check4.Value = 2
    DoEvents
    For UK& = 1 To CAFASEL&
      Y$ = REGBLAN$("LICODOSI")
      Call REPLA(Y$, Chr$(NVEN%(UK&)), 41, 1)
      Call REPLA(Y$, MKI$(NCLI%(UK&)), 19, 2)
      Call REPLA(Y$, MKI$(FEFA%(UK&)), 17, 2)
      NUDO$ = "FC": If INET#(UK&) < 0 Then NUDO$ = "NC"
      Call REPLA(Y$, NUDO$ + "-" + TRIM$(Str$(NFAC&(UK&))), 1, 16)
      Call REPLA(Y$, MKS$(NPED&(UK&)), 21, 4)
      Call REPLA(Y$, NOCO$(UK&), 25, 16)
      Call REPLA(Y$, MKD$(INET#(UK&)), 43, 8)
      POCOMIP = PORCO(NVEN%(UK&))
      If COMIP(UK&) > 0 Then POCOMIP = COMIP(UK&)
      Call REPLA(Y$, MKS$(POCOMIP), 51, 4)
      ICOM# = INET#(UK&) * POCOMIP / 100
      TOCOMI#(NVEN%(UK&)) = TOCOMI#(NVEN%(UK&)) + INET#(UK&) * POCOMIP / 100
      Call REPLA(Y$, MKD$(ICOM#), 55, 8)
      J& = J& + 1
      Call GRAREG("LICODOSI", Y$, J&)
    Next UK&
    Check4.Value = 1
    DoEvents
    Call SETULTREG("LICODOSI", J&)
    Call CIERRARCH("LICODOSI")
    '
    Call HEADERS("LICODOSI", "")
    Call HEADERS("LICODOSI", "Corresponde a: " + CORRESQ$)
    YALIQUI% = 1
    Command3.Enabled = True
    Command4.Enabled = True
    '
999 If YALIQUI% = 1 Then
        FORLICODOS.List1.Clear
        For VENDE% = 0 To 255
            If VENDE% = 0 Then
                NOVEN$ = "DIRECCION"
              Else
                NOVEN$ = ECOARCH$("VENDEDOR", Chr$(VENDE%))
            End If
            If TRIM$(NOVEN$) <> "" Then
                TXT$ = Space$(48)
                Call REPLA(TXT$, CSTRING$(MKI$(VENDE%), 1, 3, 0, 2), 1, 4)
                Call REPLA(TXT$, NOVEN$, 6, 17)
                Call REPLA(TXT$, CSTRING$(MKD$(TOCOMI#(VENDE%)), 4, 10, 2, 2), 23, 10)
                FORLICODOS.List1.AddItem TXT$
            End If
        Next VENDE%
        FORLICODOS.Check4.Value = 1
        '
        If FORLICODOS.List1.ListCount > 0 Then
            FORLICODOS.List1.ListIndex = 0
        End If
        FORLICODOS.List1.Refresh
        '
    End If
    '
    Screen.MousePointer = 1
    Exit Sub
    '
500 For UK& = 1 To CAFASEL&
      If NFAC&(UK&) = NRO1& Then
        If NCLI%(UK&) = CLI1% Then
          NREM&(UK&) = NRO2&
          GoTo 539
        End If
      End If
    Next UK&
    '
539 Return
    '
550 If RMPE& > 0 Then
      If VENPE% > 0 Then
        For UK& = 1 To CAFASEL&
          If NREM&(UK&) = RMPE& Then
            If NCLI%(UK&) = CLIPE% Then
              NPED&(UK&) = NOPE&
              NVEN%(UK&) = VENPE%
              NOCO$(UK&) = OCOPE$
              COMIP(UK&) = PORCOMIPE
              GoTo 589
            End If
          End If
        Next UK&
      End If
    End If
    '
589 Return
    '

'If NRO& = 12562 Then
'AAA = BBB
'End If
    J& = J& + 1
    Call GRAREG("LICOMIS", Y$, J&)
    TOCOMI#(VENDE%) = TOCOMI#(VENDE%) + ICOM#
    '
    If TIDOC% <> TIDOCA% Or NRO& <> NROA& Or NC% <> NCAA% Then
       MOVENTA# = 0
       ICOMIS# = 0
       JLIBRO& = JLIBRO& + 1
       TIDOCA% = TIDOC%: NROA& = NRO&: NCAA% = NC%
    End If
    '
    Y$ = REGBLAN$("LIBROVIA")
    FEX = FF%: FEY$ = FECHATEX$(FEX)
    ANIO$ = Mid$(FEY$, 7, 2)
    MESI$ = Mid$(FEY$, 4, 2)
    DIAI$ = Left$(FEY$, 2)
    If Val(ANIO$) > 55 Then
         ANIO$ = "19" + ANIO$
       Else
         ANIO$ = "20" + ANIO$
    End If
    Call REPLA(Y$, ANIO$, 1, 4)
    Call REPLA(Y$, MESI$, 5, 2)
    Call REPLA(Y$, DIAI$, 7, 2)
    '
    If TIDOC% = 8 Then
         DOCU$ = "NP-" + TRIM$(Str$(NRO&))
         SUPED#(VENDE%) = SUPED#(VENDE%) + ITOT#
         COMPED#(VENDE%) = COMPED#(VENDE%) + ICOM#
       ElseIf TIDOC% = 4 Then
         DOCU$ = "RC-" + TRIM$(Str$(NRO&))
         SUCOB#(VENDE%) = SUCOB#(VENDE%) + ITOT#
         COMCOB#(VENDE%) = COMCOB#(VENDE%) + ICOM#
       Else
         DOCU$ = FORMDOC$(NRO&, TIDOC%, VADOC%, NUSUCUVE%)
    End If
    Call REPLA(Y$, DOCU$, 9, 18)
    '
    If REGICLI&(NC%) > 0 Then
       TTTT$ = REGLEIDO$("DEUDOR1", REGICLI&(NC%))
    End If
    Call REPLA(Y$, RASOCLI$(NC%), 27, 30)
    Call REPLA(Y$, "Acep.", 57, 12)
    Call REPLA(Y$, MKI$(FF%), 69, 2)
    '
    MOVENTA# = MOVENTA# + ITOT#
    ICOMIS# = ICOMIS# + ICOM#
    PORCOMIS = PORCOM
    If Abs(MOVENTA#) > 0.05 Then
       If PORCOMIS < 0.1 Then
          PORCOMIS = 100 * ICOMIS# / MOVENTA#
       End If
    End If
    '
    Call REPLA(Y$, MKS$(PORCOMIS), 71, 4)
    Call REPLA(Y$, MKD$(MOVENTA#), 75, 8)
    Call REPLA(Y$, MKD$(ICOMIS#), 83, 8)
    '
    OBSER$ = Space$(38)
    Call REPLA(Y$, OBSER$, 91, 36)
    Call REPLA(Y$, MKI$(VENDE%), 127, 2)
    '
'If NRO& = 12562 Then
'AAA = CVI(Mid$(Y$, 127, 2))
'BBB = VENDE%
'End If
    Call GRAREG("LIBROVIA", Y$, JLIBRO&)
    '
Return
    '
    '
610 For I& = J& + 1 To ULTREG&("LICOMIS")
        Call GRAREG("LICOMIS", String$(64, 0), I&)
    Next I&
    Call CIERRARCH("LICOMIS")
    '
    Call SETULTREG("LIBROVIA", JLIBRO&)
    Call CIERRARCH("LIBROVIA")
    '
Return
'
'
End Sub
'


