VERSION 5.00
Begin VB.Form FLICOMIS 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Liquidación de Comisiones por Venta y Cobranza"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10920
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   10920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   9870
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   24
      Top             =   3000
      Width           =   855
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   25
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Modalidad de Cálculo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      Left            =   210
      TabIndex        =   19
      Top             =   1780
      Width           =   4635
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Por Cobranza Neta de I.V.A"
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
         TabIndex        =   23
         Top             =   1575
         Width           =   3375
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Por Facturación Neta del Periodo"
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
         TabIndex        =   22
         Top             =   1260
         Value           =   -1  'True
         Width           =   3375
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Por Pedidos Netos y Cobranza"
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
         TabIndex        =   21
         Top             =   945
         Width           =   4005
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Por Cobranza Neta según Artículo"
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
         TabIndex        =   20
         Top             =   420
         Width           =   3375
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Imprimir"
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
      Left            =   9870
      Picture         =   "FLICOMIS.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1885
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
      Height          =   1740
      Left            =   5040
      TabIndex        =   16
      Top             =   2060
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
         Height          =   1185
         Left            =   210
         TabIndex        =   17
         Top             =   380
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
      Top             =   320
      Width           =   4635
      Begin VB.TextBox MESAN 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1575
         TabIndex        =   10
         Text            =   " "
         Top             =   210
         Width           =   2220
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   3045
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   " "
         Top             =   735
         Width           =   750
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   1575
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   " "
         Top             =   735
         Width           =   1065
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         Left            =   3885
         TabIndex        =   13
         Top             =   840
         Width           =   1590
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   225
         Left            =   2835
         TabIndex        =   12
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
         TabIndex        =   11
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
      Height          =   1725
      Left            =   5040
      TabIndex        =   2
      Top             =   210
      Width           =   4635
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
         Caption         =   "Valorización por Producto"
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
         Width           =   2745
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Detalle de Facturas por Recibo"
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
         Caption         =   "Selección de Cobranza del Período"
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
      Height          =   750
      Left            =   9870
      Picture         =   "FLICOMIS.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1100
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
      Height          =   750
      Left            =   9870
      Picture         =   "FLICOMIS.frx":0974
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   315
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   9260
      Picture         =   "FLICOMIS.frx":0ABE
      Top             =   3260
      Width           =   2010
   End
End
Attribute VB_Name = "FLICOMIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PORCO(256)
Dim SUPED#(256), COMPED#(256), SUCOB#(256), COMCOB#(256), TOCOMI#(256), TOCOB#(256)
Dim NRECI&(8192), REGFA$(8192)
Dim NCV%(8192), NROV&(8192), IDET#(8192), IFAC#(8192)


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
    Command2.Enabled = False
    Command3.Enabled = False
    YALIQUI% = 0
    MODOCAL% = 0
    If Option2.Value = True Then MODOCAL% = 1
    If Option3.Value = True Then MODOCAL% = 2
    If Option4.Value = True Then MODOCAL% = 3
    If Option5.Value = True Then MODOCAL% = 4
    If MODOCAL% < 1 Then
        Call MENSERR(24, "Falta Seleccionar Modalidad de Cálculo")
        GoTo 99
    End If
    If MODOCAL% = 1 Then
        If InStr(UCase$(EMPREAC$), "NOVOESPU") > 0 Then
            Call CALCOMIS
          Else
            Call OPNOIN("")
        End If
        GoTo 99
    End If
    '
    If MODOCAL% > 1 Then
'        If InStr(UCase$(EMPREAC$), "SABO") > 0 Then
            Call LICOMIS
'          Else
'            Call OPNOIN("")
'        End If
        GoTo 99
    End If
    '
99  Command2.Enabled = True
    If YALIQUI% = 1 Then
        Command3.Enabled = True
    End If
    '
End Sub

Private Sub Command3_Click()
    Call PRILICOMIS
End Sub

Private Sub Form_Activate()
    If List1.ListCount < 1 Then
        Command3.Enabled = False
    End If
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
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CIERRARCH("*.*")
End Sub

Private Sub List1_DblClick()
    Call PRILICOMIS
End Sub

Private Sub MESAN_GOTFOCUS()
    Command2.SetFocus
End Sub

Private Sub Option2_Click()
    List1.Clear
    Command3.Enabled = False
End Sub

Private Sub Option3_Click()
    List1.Clear
    Command3.Enabled = False
End Sub

Private Sub Option4_Click()
    List1.Clear
    Command3.Enabled = False
End Sub

Private Sub Option5_Click()
    List1.Clear
    Command3.Enabled = False
End Sub

Private Sub Text4_GOTFOCUS(Index As Integer)
    Command2.SetFocus
End Sub

Sub PAGOVEN()
   '
5  VDEF$ = REGBLAN$("CUECOMIS")
   HOII% = HOY(HOS$)
   Call REPLA(VDEF$, MKI$(HOII%), 2, 2)
   Call REPLA(VDEF$, Chr$(2), 4, 1)
   '
10 OBX$ = OBJPLA$("PAGOCOMIS", VDEF$)
   If OBX$ <> "" Then
     VENDE% = Asc(Left$(OBX$, 1))
     If VENDE% < 1 Or ECOARCH("VENDEDOR", Chr$(VENDE%)) = "" Then
       VDEF$ = OBX$
       Call MENSERR(24, "Falta Vendedor")
       GoTo 10
     End If
     '
     REFE$ = TRIM$(Mid$(OBX$, 9, 32))
     If REFE$ = "" Then
       VDEF$ = OBX$
       Call MENSERR(24, "Falta Referencia")
       GoTo 10
     End If
     '
     IMPO# = CVD(Mid$(OBX$, 49, 8))
     If Abs(IMPO#) < 0.005 Then
       VDEF$ = OBX$
       Call MENSERR(24, "Falta Importe")
       GoTo 10
     End If
     '
     If IMPO# < 0 Then
       Call REPLA(VDEF$, MKD$(Abs(IMPO#)), 41, 8)
       Call REPLA(VDEF$, MKD$(0), 49, 8)
     End If
     '
     Call REGAPP("CUECOMIS", OBX$)
     Call CIERRARCH("CUECOMIS")
     '
     Call COMUNI("Pago Registrado !!!")
     GoTo 5
     '
   End If
   '
End Sub

Sub CALCOMIS()
    '
    For U& = 1 To ULTREG&("VENDEDOR")
        X$ = REGLEIDO$("VENDEDOR", U&)
        VENDE% = Asc(Left$(X$, 1))
        If VENDE% > 0 Then
            If VENDE% <= 256 Then
                If PORCO(VENDE%) < 0.005 Then
                    PORCO(VENDE%) = CVS(Mid$(X$, 44, 4)) + CVS(Mid$(X$, 48, 4))
                End If
            End If
        End If
    Next U&
    Call CIERRARCH("VENDEDOR")
    '
    Call DESHASFECHA(DES%, HAS%, PERIO$)
    If PERIO$ = "" Then Exit Sub
    MESAN = PERIO$
    '
    For U& = 1 To ULTREG&("DETCOBV")
        X$ = REGLEIDO$("DETCOBV", U&)
        NRECI&(U&) = CVS(Mid$(X$, 5, 4))
        REGFA$(U&) = ""
    Next U&
    '
    CARECI& = U&
    '
    FIN& = ULTREG&("APLICOB")
    FLICOMIS.Check1.Value = 2
    '
    FIN& = ULTREG("APLICOB")
    For U& = 1 To FIN&
        Call TRACE(20, U&, FIN&)
        X$ = REGLEIDO$("APLICOB", U&)
        FEDOC% = CVI(Left$(X$, 2))
        If FEDOC% >= DES% Then
           If FEDOC% <= HAS% Then
              NRO& = Val(Mid$(X$, 15, 8))
              If NRO& > 0 Then
                 If Mid$(X$, 55, 2) = "FC" Then
                    For UU& = 1 To CARECI&
                       If NRECI&(UU&) = NRO& Then
                          REGFA$(UU&) = REGFA$(UU&) + MKS$(U&)
                          GoTo 19
                       End If
                    Next UU&
                 End If
              End If
           End If
        End If
19  Next U&
    FLICOMIS.Check1.Value = 1
    '
    FIN& = CARECI&
    FLICOMIS.Check2.Value = 2
    '
    CAFAC% = 0
    Call BLANARCH("LICOMIN")
    For U& = 1 To CARECI&
        Call TRACE(20, U&, CARECI&)
        X$ = REGLEIDO$("DETCOBV", U&)
        NRO& = CVS(Mid$(X$, 5, 4))
        FEDOC% = CVI(Mid$(X$, 3, 2))
        NC% = CVI(Mid$(X$, 9, 2))
        ICOB# = CVD(Mid$(X$, 11, 8))
        INOAPLI# = ICOB#
        '
        For V& = 1 To Len(REGFA$(U&)) Step 4
            W& = CVS(Mid$(REGFA$(U&), V&, 4))
            If W& > 0 Then
               If W& <= ULTREG&("APLICOB") Then
                  Y$ = REGLEIDO$("APLICOB", W&)
                  IAPLI# = CVD(Mid$(Y$, 23, 8))
                  If IAPLI# > 0 Then
                     NROFAC& = Val(Mid$(Y$, 63, 8))
                     INOAPLI# = INOAPLI# - IAPLI#
                     '
                     Z$ = REGBLAN$("LICOMIN")
                     Call REPLA(Z$, MKI$(NC%), 1, 2)
                     Call REPLA(Z$, MKS$(NRO&), 3, 4)
                     Call REPLA(Z$, MKI$(FEDOC%), 7, 2)
                     Call REPLA(Z$, MKD$(ICOB#), 9, 8)
                     Call REPLA(Z$, MKI$(NC%), 17, 2)
                     Call REPLA(Z$, Chr$(VENDCLI%(NC%)), 19, 2)
                     '
                     Call REPLA(Z$, MKS$(NROFAC&), 21, 4)
                     Call REPLA(Z$, MKD$(IAPLI#), 25, 8)
                     Call REGAPP("LICOMIN", Z$)
                     '
                     CAFAC% = CAFAC% + 1
                     NCV%(CAFAC%) = NC%
                     NROV&(CAFAC%) = NROFAC&
                     IDET#(CAFAC%) = 0
                     IFAC#(CAFAC%) = 0
                  End If
               End If
            End If
        Next V&
        '
        If Abs(INOAPLI#) >= 0.005 Then
            Z$ = REGBLAN$("LICOMIN")
            Call REPLA(Z$, MKI$(NC%), 1, 2)
            Call REPLA(Z$, MKS$(NRO&), 3, 4)
            Call REPLA(Z$, MKI$(FEDOC%), 7, 2)
            Call REPLA(Z$, MKD$(ICOB#), 9, 8)
            Call REPLA(Z$, MKI$(NC%), 17, 2)
            Call REPLA(Z$, Chr$(VENDCLI%(NC%)), 19, 2)
            '
            Call REPLA(Z$, MKS$(0), 21, 4)
            Call REPLA(Z$, MKD$(INOAPLI#), 25, 8)
            Call REGAPP("LICOMIN", Z$)
            '
            CAFAC% = CAFAC% + 1
            NCV%(CAFAC%) = NC%
            NROV&(CAFAC%) = 0
            IDET#(CAFAC%) = 0
            IFAC#(CAFAC%) = 0
            '
        End If
        '
    Next U&
    FLICOMIS.Check2.Value = 1
    '
    FIN& = ULTREG&("MOVISTO")
    FLICOMIS.Check3.Value = 2
    '
    For IL& = 1 To FIN&
        Call TRACE(20, IL&, FIN&)
        X$ = REGLEIDO$("MOVISTO", IL&)
        COMO$ = UCase$(Mid$(X$, 33, 2))
        If COMO$ = "FC" Then
          NRO1& = Val(Mid$(X$, 36, 7))
          If NRO1& > 0 Then
             NC1% = CVI(Mid$(X$, 84, 2))
             For KI% = 1 To CAFAC%
                If NROV&(KI%) = NRO1& Then
                   If NCV%(KI%) = NC1% Then
                      CAN# = CVD(Mid$(X$, 96, 8)) - CVD(Mid$(X$, 88, 8))
                      PRE# = CVD(Mid$(X$, 73, 8))
                      TOT# = PRE# * CAN#
                      CI% = CVI(Mid$(X$, 3, 2))
                      FACAPRE% = 100
                      '
                      If CI% > 0 Then
                         If CI% <= NUMAS% Then
                            REGMA& = CI% + 1
                            Y$ = REGLEIDO$("MASTER", REGMA&)
                            FACAPRE% = CVI(Mid$(Y$, 127, 2))
                            If FACAPRE% < 1 Then
30                             VDEF$ = MKI$(CI%) + MKI$(100)
                               If FACAPRE% >= 1 Then
                                   Call REPLA(VDEF$, MKI$(FACAPRE%), 3, 2)
                               End If
                               OBX$ = OBJPLA$("FACAPRE", VDEF$)
                               If OBX$ = "" Then
                                  Call CIERRARCH("*.*")
                                  Exit Sub
                               End If
                               FACAPRE1% = CVI(Mid$(OBX$, 3, 2))
                               If FACAPRE1% < 1 Or FACAPRE1% > 32767 Then GoTo 30
                               If REGMA& > 1 Then
                                  Y$ = REGLEIDO$("MASTER", REGMA&)
                                  Call REPLA(Y$, MKI$(FACAPRE1%), 127, 2)
                                  Call GRAREG("MASTER", Y$, REGMA&)
                               End If
                               FACAPRE% = FACAPRE1%
                            End If
                         End If
                      End If
                      '
                      TOPONDE# = TOT# * FACAPRE% / 100
                      '
                      IDET#(KI%) = IDET#(KI%) + TOT#
                      IFAC#(KI%) = IFAC#(KI%) + TOPONDE#
                      '
                   End If
               End If
             Next KI%
          End If
          '
        End If
    Next IL&
    FLICOMIS.Check3.Value = 1
    '
    FIN& = CAFAC%
    FLICOMIS.Check4.Value = 2
    '
    FIN& = CAFAC%
    For IL& = 1 To CAFAC%
        Call TRACE(20, IL&, FIN&)
        X$ = REGLEIDO$("LICOMIN", IL&)
        COEF = 1
        If IDET#(IL&) > 0 Then
            COEF = IFAC#(IL&) / IDET#(IL&)
        End If
        Call REPLA(X$, MKS$(COEF), 41, 4)
        '
        VENDE% = Asc(Mid$(X$, 19, 1))
        If VENDE% >= 0 Then
           If VENDE% <= 256 Then
              Call REPLA(X$, MKS$(PORCO(VENDE%)), 45, 4)
              ICOMI# = (Int(CVD(Mid$(X$, 25, 8)) * COEF * PORCO(VENDE%) + 0.5)) / 100
              Call REPLA(X$, MKD$(ICOMI#), 49, 8)
              TOCOMI#(VENDE%) = TOCOMI#(VENDE%) + ICOMI#
              TOCOB#(VENDE%) = TOCOB#(VENDE%) + CVD(Mid$(X$, 25, 8))
           End If
        End If
        '
        Call GRAREG("LICOMIN", X$, IL&)
        '
    Next IL&
    '
    Call CIERRARCH("LICOMIN")
    FLICOMIS.List1.Clear
    For VENDE% = 0 To 255
      If TOCOB#(VENDE%) > 0 Then
        If VENDE% = 0 Then
             NOVEN$ = "DIRECCION"
           Else
             NOVEN$ = ECOARCH$("VENDEDOR", Chr$(VENDE%))
        End If
        TXT$ = Space$(48)
        Call REPLA(TXT$, CSTRING$(MKI$(VENDE%), 1, 3, 0, 2), 1, 4)
        Call REPLA(TXT$, NOVEN$, 6, 17)
        Call REPLA(TXT$, CSTRING$(MKD$(TOCOMI#(VENDE%)), 4, 10, 2, 2), 23, 10)
        FLICOMIS.List1.AddItem TXT$
      End If
    Next VENDE%
    FLICOMIS.Check4.Value = 1
    '
    If FLICOMIS.List1.ListCount > 0 Then
      FLICOMIS.List1.ListIndex = 0
    End If
    FLICOMIS.List1.Refresh
    FLICOMIS.Command3.Enabled = True
    '
End Sub
'

Sub LICOMIS()
    '
    Screen.MousePointer = 11
    '
    YALIQUI% = 0: CUECOMI% = 0
    Dim CVX%(128), COVEN(128), COCOB(128)
    Dim PORCO(256), TOCOMI#(256), TOCOB#(256)
    Dim REGICLI&(32767), VENDECLI%(32767)
    '
    If EXISTE%(LUDAT$ + "CUECOMIS.RFS") Then
      CUECOMI% = 1
      Screen.MousePointer = 11
      Call BLOQARCH("CUECOMIS")
      FIN& = ULTREG&("CUECOMIS")
      JJ& = 0
      For U& = 1 To FIN&
        XX$ = REGLEIDO$("CUECOMIS", U&)
        FECO% = CVI(Mid$(XX$, 2, 2))
        TIMX% = Asc(Mid$(XX$, 4, 1))
        If FECO% < DES% Or FECO% > HAS% Or TIMX% = 2 Then
          JJ& = JJ& + 1
          Call GRAREG("CUECOMIS", XX$, JJ&)
        End If
      Next U&
      Call SETULTREG("CUECOMIS", JJ&)
      Call CIERRARCH("CUECOMIS")
      Screen.MousePointer = 1
    End If
    '
    For I& = 1 To 256
       SUPED#(I&) = 0
       COMPED#(I&) = 0
       SUCOB#(I&) = 0
       COMCOB#(I&) = 0
    Next I&
    '
    NUVE% = 0
    YAGERE% = 0
    For I& = 1 To ULTREG&("VENDEDOR")
        X$ = REGLEIDO$("VENDEDOR", I&)
        CV% = Asc(Left$(X$, 1))
        If CV% > 0 Then
            NUVE% = NUVE% + 1
            CVX%(NUVE%) = CV%
            COVEN(NUVE%) = CVS(Mid$(X$, 44, 4))
            COCOB(NUVE%) = CVS(Mid$(X$, 48, 4))
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
        COCOB(NUVE%) = CVS(Mid$(X$, 48, 4))
    End If
    Call CIERRARCH("VENDEDOR")
    '
50  Call DESHASFECHA(DES%, HAS%, PERIO$)
    If PERIO$ = "" Then Exit Sub
    MESAN = PERIO$
    '
    Call HEADERS("LICOMIX", "")
    Call HEADERS("LICOMIX", "Corresponde a: " + CORRESP$)
    '
    FIN& = ULTREG&("DEUDOR1")
    URE2& = ULTREG&("DEUDOR2")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("DEUDOR1", U&)
      If U& <= URE2& Then
        X$ = X$ + REGLEIDO$("DEUDOR2", U&)
        NC% = CVI(Left$(X$, 2))
        If NC% > 0 Then
          If NC% <= 32767 Then
            If REGICLI&(NC%) < 1 Then
              REGICLI&(NC%) = U&
              VENDECLI%(NC%) = Asc(Mid$(X$, 152, 1))
            End If
          End If
        End If
      End If
    Next U&
    '
    'For U& = 1 To ULTREG&("HISTOVEN")
    '  X$ = REGLEIDO$("HISTOVEN", U&)
    '  NC% = CVI(Left$(X$, 2))
    '  VENDECLI%(NC%) = CVI(Mid$(X$, 3, 2))
    'Next U&
    'Call CIERRARCH("HISTOVEN")
    '
    MODO% = 0
    If FLICOMIS.Option3.Value = True Then MODO% = 1
    If FLICOMIS.Option4.Value = True Then MODO% = 2
    If FLICOMIS.Option5.Value = True Then MODO% = 3
    '
    J& = 0: JLIBRO& = 0
    On MODO% GoTo 100, 200, 300
    GoTo 999
    '
100 Call HEADERS("LICOMIX", "Por: Pedidos")
    '
    FIN& = ULTREG&("PEDDETA")
    '
    INIPE& = 0
    For I& = 1 To FIN&
        Call TRACE(20, I& - INIPE&, FIN& - INIPE&)
        X$ = REGLEIDO$("PEDDETA", I&)
        FF% = CVI(Mid$(X$, 5, 2))
        If FF% >= DES% Then
            If FF% <= HAS% Then
                NRO& = CVS(Left$(X$, 4))
                If NRO& > 0 Then
If NRO& >= 44357 And NRO& <= 44370 Then GoTo 129
If NRO& >= 44428 And NRO& <= 44435 Then GoTo 129
If NRO& = 44440 Then GoTo 129
                    NC% = CVI(Mid$(X$, 7, 2))
                    If NC% > 0 Then
                       VENDE% = 0 ' VENDE%=ASC(MID$(X$,30,1))
                       If VENDE% < 1 Then VENDE% = VENDECLI%(NC%)
                       If VENDE% < 1 Then VENDE% = 99
                       '
                       CI% = CVI(Mid$(X$, 33, 2))
                       CAN = CVS(Mid$(X$, 79, 4))
                       If REGICLI&(NC%) > 0 Then
                          TTTT$ = REGLEIDO$("DEUDOR1", REGICLI&(NC%))
                       End If
                       DESCU = PORDESCLI(NC%)
                       PREUNI = CVS(Mid$(X$, 83, 4)) * (1 - DESCU / 100)
                       ITOT# = (Int(100 * CDbl(PREUNI) * CAN + 0.5)) / 100
                       TIDOC% = 8
                       '
                       For K% = 1 To NUVE%
                          If CVX%(K%) = VENDE% Then
                             PORCOM = COVEN(K%)
' puesto 3/8/00 para reiba por pedido de sabo - provisorio
                             If NC% = 1271 Then PORCOM = 1
                             GoTo 120
                          End If
                       Next K%
                       PORCOM = 0
                       '
120                    ICOM# = (Int(ITOT# * PORCOM + 0.5)) / 100
                       NUSUCUVE% = 0
                       GoSub 600 '
                       If INIPE& = 0 Then INIPE& = I&
                       '
                    End If
                End If
            End If
        End If
129 Next I&
    '
    GoTo 300                 ' PARA INCLUIR COBRANZA
    GoSub 610
    Call FILESORT("LICOMIS", "", 3, 5)
    YALIQUI% = 1
    GoTo 999
    '
    '
200 Call HEADERS("LICOMIX", "Por: Facturacion")
    FIN& = ULTREG&("FACTUR")
    '
    For I& = 1 To FIN&
        Call TRACE(20, I&, FIN&)
        X$ = REGLEIDO$("FACTUR", I&)
        FF% = CVI(Mid$(X$, 7, 2))
        If FF% >= DES% Then
            If FF% <= HAS% Then
                NRO& = CVS(Mid$(X$, 3, 4))
                If NRO& > 0 Then
                    NC% = CVI(Mid$(X$, 1, 2))
                    If NC% > 0 Then
                       VENDE% = VENDECLI%(NC%)
                       If VENDE% < 1 Then VENDE% = 99
                       '
                       CI% = 0
                       CAN = 0
                       DESCU = 0
                       PREUNI = 0
                       ITOT# = CVD(Mid$(X$, 15, 8)) + CVD(Mid$(X$, 49, 8))
                       TIDOC% = Asc(Mid$(X$, 13, 1))
                       VADOC% = Asc(Mid$(X$, 14, 1))
                       If TIDOC% >= 4 Then
                           ITOT# = (-1) * Abs(ITOT#)
                       End If
                       '
                       NUSUCUVE% = CVI(Mid$(X$, 47, 2))
                       If TIDOC% > 0 Then
                          For K% = 1 To NUVE%
                             If CVX%(K%) = VENDE% Then
                                PORCOM = COVEN(K%)
                                GoTo 220
                             End If
                          Next K%
                          PORCOM = 0
                          '
220                       ICOM# = (Int(ITOT# * PORCOM + 0.5)) / 100
                          GoSub 600
                          '
                       End If
                    End If
                End If
            End If
        End If
    Next I&
    '
    GoSub 610
    Call FILESORT("LICOMIS", "", 3, 5)
    YALIQUI% = 1
    GoTo 999
    '
    '
300 Call HEADERS("LICOMIX", "Por: Cobranza")
    FIN& = ULTREG&("COBRAN")
    '
    For I& = 1 To FIN&
        Call TRACE(20, I&, FIN&)
        X$ = REGLEIDO$("COBRAN", I&)
        FF% = CVI(Mid$(X$, 7, 2))
        If FF% >= DES% Then
            If FF% <= HAS% Then
                NRO& = CVS(Mid$(X$, 3, 4))
                If NRO& > 0 Then
                    NC% = CVI(Mid$(X$, 1, 2))
                    If NC% > 0 Then
                       VENDE% = VENDECLI%(NC%)
                       If VENDE% < 1 Then VENDE% = 99
                       '
                       CI% = 0
                       CAN = 0
                       DESCU = 0
                       PREUNI = 0
                       ITOT# = CVD(Mid$(X$, 15, 8))
                       TIDOC% = 4
                       '
                       NUSUCUVE% = 0
                       If TIDOC% > 0 Then
                          For K% = 1 To NUVE%
                             If CVX%(K%) = VENDE% Then
                                PORCOM = COCOB(K%)
                                GoTo 320
                             End If
                          Next K%
                          PORCOM = 0
                          '
320                       ' NO CONSIDERAR EL IVA
                          PVCL% = PIVACLI%(NC%)
                          DIVI = 1 + ALIVA / 100
                          If PVCL% = 2 Then DIVI = 1 + 1.5 * ALIVA / 100
                          If PVCL% = 5 Or PVCL% = 7 Then DIVI = 1
                          ICOM# = (Int(ITOT# * PORCOM / DIVI + 0.5)) / 100
                          GoSub 600
                          '
                       End If
                    End If
                End If
            End If
        End If
    Next I&
    '
    GoSub 610
    Call FILESORT("LICOMIS", "", 3, 5)
    YALIQUI% = 1
    GoTo 999
    '
    '
999 If YALIQUI% = 1 Then
        FLICOMIS.List1.Clear
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
                FLICOMIS.List1.AddItem TXT$
            End If
        Next VENDE%
        FLICOMIS.Check4.Value = 1
        '
        If FLICOMIS.List1.ListCount > 0 Then
            FLICOMIS.List1.ListIndex = 0
        End If
        FLICOMIS.List1.Refresh
        '
    End If
    '
    Screen.MousePointer = 1
    Exit Sub
    '
    '
600 Y$ = REGBLAN$("LICOMIS")
    Call REPLA(Y$, Chr$(VENDE%), 1, 1)
    Call REPLA(Y$, MKI$(NC%), 2, 2)
    Call REPLA(Y$, MKI$(FF%), 4, 2)
    Call REPLA(Y$, Chr$(TIDOC%), 6, 1)
    Call REPLA(Y$, MKS$(NRO&), 7, 4)
    Call REPLA(Y$, MKI$(CI%), 11, 2)
    Call REPLA(Y$, MKS$(PREUNI), 13, 4)
    Call REPLA(Y$, MKS$(CAN), 17, 4)
    Call REPLA(Y$, MKD$(ITOT#), 21, 8)
    Call REPLA(Y$, MKS$(PORCOM), 29, 4)
    Call REPLA(Y$, MKD$(ICOM#), 33, 8)
    '
'If NRO& = 12562 Then
'AAA = BBB
'End If
    J& = J& + 1
    Call GRAREG("LICOMIS", Y$, J&)
    TOCOMI#(VENDE%) = TOCOMI#(VENDE%) + ICOM#
    '
    If CUECOMI% = 1 Then
      YCM$ = REGBLAN$("CUECOMIS")
      Call REPLA(YCM$, Chr$(VENDE%), 1, 1)
      Call REPLA(YCM$, MKI$(FF%), 2, 2)
      Call REPLA(YCM$, Chr$(1), 4, 1)
      Call REPLA(YCM$, MKS$(NRO&), 5, 4)
      Call REPLA(YCM$, RASOCLI$(NC%), 9, 32)
      Call REPLA(YCM$, MKD$(ICOM#), 41, 8)
      Call REGAPP("CUECOMIS", YCM$)
    End If
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
    Call GRAREG("LIBROVIA", Y$, JLIBRO&)
    '
Return
    '
    '
610 For I& = J& + 1 To ULTREG&("LICOMIS")
        Call GRAREG("LICOMIS", String$(64, 0), I&)
    Next I&
    Call CIERRARCH("LICOMIS")
    If CUECOMI% = 1 Then
      Call CIERRARCH("CUECOMIS")
    End If
    '
    Call SETULTREG("LIBROVIA", JLIBRO&)
    Call CIERRARCH("LIBROVIA")
    '
Return
'
'
End Sub
'

Sub PRILICOMIS()
    '
    MODOCAL% = 0
    If FLICOMIS.Option2.Value = True Then MODOCAL% = 1
    If FLICOMIS.Option3.Value = True Then MODOCAL% = 2: MODOCA$ = "Por Pedidos Netos de Descuento"
    If FLICOMIS.Option4.Value = True Then MODOCAL% = 3: MODOCA$ = "Por Facturacion"
    If FLICOMIS.Option5.Value = True Then MODOCAL% = 4: MODOCA$ = "Por Cobranza Neta de IVA"
    If MODOCAL% < 1 Then
        Exit Sub
    End If
    '
    If FLICOMIS.List1.ListIndex >= 0 Then
        TXT$ = FLICOMIS.List1.TEXT
        VENDE% = Val(Left$(TXT$, 3))
        NOVEN$ = "Clientes de Direccion"
        If VENDE% > 0 Then
            NOVEN$ = ECOARCH$("VENDEDOR", Chr$(VENDE%))
        End If
        '
        CORRESP$ = MESAN.TEXT
        '
        If MODOCAL% = 1 Then
            Call COPYSTRU("LICOMIN", "LICOMIV")
            Call BLANARCH("LICOMIV")
            Call HEADERS("LICOMIV", "")
            Call HEADERS("LICOMIV", "Corresponde a: " + CORRESP$)
            Call HEADERS("LICOMIV", "Vendedor: " + NOVEN$)
            '
            FIN& = ULTREG&("LICOMIN")
            '
            For IL& = 1 To FIN&
                Call TRACE(20, IL&, FIN&)
                X$ = REGLEIDO$("LICOMIN", IL&)
                If Asc(Mid$(X$, 19, 1)) = VENDE% Then
                    Call REGAPP("LICOMIV", X$)
                End If
            Next IL&
            '
            Call CIERRARCH("*.*")
            Call CONECRUN("*LICOMIV", "Listado de Comisiones Liquidadas")
            Exit Sub
            '
          Else
            J& = 0
            Call HEADERS("LICOMIX", "")
            
            Call HEADERS("LICOMIX", "Corresponde a: " + CORRESP$)
            Call HEADERS("LICOMIX", "Modalidad: " + MODOCA$)
            Call HEADERS("LICOMIX", "Vendedor: " + NOVEN$)
            '
            FIN& = ULTREG&("LICOMIS")
            '
            For IL& = 1 To FIN&
                Call TRACE(20, IL&, FIN&)
                X$ = REGLEIDO$("LICOMIS", IL&)
                If Asc(Mid$(X$, 1, 1)) = VENDE% Then
                    J& = J& + 1
                    Call GRAREG("LICOMIX", X$, J&)
                End If
            Next IL&
            '
            Call SETULTREG("LICOMIX", J&)
            Call CIERRARCH("*.*")
            Call CONECRUN("*LISLICO", "Listado de Comisiones Liquidadas")
            Exit Sub
        End If
        '
      Else
        '
        Call MENSERR(24, "Falta Seleccionar Vendedor")
        '
    End If
    '
End Sub
'

