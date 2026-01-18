VERSION 5.00
Begin VB.Form FORLICO 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Liquidación de Comisiones por Venta y Cobranza"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10920
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   10920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Libro"
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
      Left            =   9870
      Picture         =   "FORLICO.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Imprime Libro de Viajantes"
      Top             =   3150
      Width           =   855
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFF80&
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
      TabIndex        =   21
      Top             =   1680
      Width           =   4635
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FFFF80&
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
         TabIndex        =   25
         Top             =   1575
         Width           =   3375
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FFFF80&
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
         TabIndex        =   24
         Top             =   1260
         Width           =   3375
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFFF80&
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
         TabIndex        =   23
         Top             =   945
         Value           =   -1  'True
         Width           =   4005
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFF80&
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
         TabIndex        =   22
         Top             =   420
         Width           =   3375
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFF80&
      Caption         =   "Opciones de Cálculo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   210
      TabIndex        =   19
      Top             =   3885
      Width           =   4635
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFF80&
         Caption         =   "Presenta F.C.P. por cada Producto"
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
      Height          =   855
      Left            =   9870
      Picture         =   "FORLICO.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2205
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF80&
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
      Height          =   2115
      Left            =   5040
      TabIndex        =   16
      Top             =   2625
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
         Height          =   1635
         Left            =   210
         TabIndex        =   17
         Top             =   315
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
      Begin VB.TextBox MESAN 
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
      BackColor       =   &H00FFFF80&
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
      Left            =   5040
      TabIndex        =   2
      Top             =   210
      Width           =   4635
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   525
         ScaleHeight     =   75
         ScaleWidth      =   3525
         TabIndex        =   26
         Top             =   1785
         Width           =   3585
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   27
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00FFFF80&
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
         BackColor       =   &H00FFFF80&
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
         BackColor       =   &H00FFFF80&
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
         BackColor       =   &H00FFFF80&
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
      Height          =   855
      Left            =   9870
      Picture         =   "FORLICO.frx":0974
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1260
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
      Left            =   9870
      Picture         =   "FORLICO.frx":0C7E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   315
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   555
      Left            =   9620
      Picture         =   "FORLICO.frx":0DC8
      Stretch         =   -1  'True
      Top             =   4170
      Width           =   5250
   End
End
Attribute VB_Name = "FORLICO"
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
    FORLICO.Hide
    Close: End
End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    Command3.Enabled = False
    '
    Call LICOMIS
    '
99  Command2.Enabled = True
    Command3.Enabled = True
    Command4.Enabled = True
    '
End Sub

Private Sub Command3_Click()
    Call PRILICOMIS
End Sub

Private Sub Command4_Click()
    Call PRILIVIAJAN
End Sub

Private Sub Form_Load()
    '
    Call VERJOBID
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
    AMODO% = ALTERNA%("Listado de Comisiones\Libro de Viajantes")
    If AMODO% = 1 Then
         Call PRILICOMIS
       ElseIf AMODO% = 2 Then
         Call PRILIVIAJAN
       Else
         Exit Sub
    End If
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
Private Sub Command22_Click()
Dim DEVEN$(3)
DEVEN$(1) = "MAESTRI, LUIS"
DEVEN$(2) = "LENOBEL, ALFREDO"
DEVEN$(3) = "MOREIRA, CLAUDIO F."
'
   Call CARSELMESA
10 Call SELECHO("SELMESA")
   MESASEL$ = VALACT1$("SELMESA")
   If TRIM$(MESASEL$) <> "" Then
      ABC$ = MESANO$(MESASEL$, DES%, HAS%)
      DD1% = DIPOST(HAS%)
      BCD$ = METEXTO$(DD1%)
      ARCHICOX$ = "LCS" + Left$(BCD$, 3) + Right$(BCD$, 2)
      Call COPYSTRU("LICOMIS", ARCHICOX$)
      '
      For VENDE% = 1 To 3
        J& = 0
        Call HEADERS("LICOMIX", "")
          
        Call HEADERS("LICOMIX", "Corresponde a: " + MESASEL$)
        Call HEADERS("LICOMIX", "Vendedor: " + DEVEN$(VENDE%))
        '
        FIN& = ULTREG&(ARCHICOX$)
        '
        For IL& = 1 To FIN&
          Call TRACE(20, IL&, FIN&)
          X$ = REGLEIDO$(ARCHICOX$, IL&)
          If Asc(Mid$(X$, 1, 1)) = VENDE% Then
            J& = J& + 1
            Call GRAREG("LICOMIX", X$, J&)
          End If
        Next IL&
        '
        Call SETULTREG("LICOMIX", J&)
        Call CIERRARCH("*.*")
        Call CONECRUN("*LISLICO", "Listado de Comisiones Liquidadas")
      Next VENDE%
      GoTo 10
   End If
   '
End Sub

Private Sub Command23_Click()
   Dim CMES$(13)
   CMES$(1) = "NOV00"
   CMES$(2) = "DIC00"
   CMES$(3) = "ENE01"
   CMES$(4) = "FEB01"
   CMES$(5) = "MAR01"
   CMES$(6) = "ABR01"
   CMES$(7) = "MAY01"
   CMES$(8) = "JUN01"
   CMES$(9) = "JUL01"
   CMES$(10) = "AGO01"
   CMES$(11) = "SET01"
   CMES$(12) = "OCT01"
   CMES$(13) = "NOV01"
   '
   NCLIA% = 0
   JJ& = 0
   For UI% = 1 To 13
     ABRIX$ = "LCS" + CMES$(UI%)
     Call COPYSTRU("LICOMIS", ABRIX$)
     For UJ& = 1 To ULTREG&(ABRIX$)
       X$ = REGLEIDO$(ABRIX$, UJ&)
       VENX% = Asc(Left$(X$, 1))
       NCLIE% = CVI(Mid$(X$, 2, 2))
       VCLX% = VENDCLI%(NCLIE%)
       If VCLX% < 1 Then VCLX% = 99
       If VENX% <> VCLX% Then
         If NCLIE% <> NCLIA% Then
           NCLIA% = NCLIE%
           Z$ = MKI$(NCLIA%) + MKI$(VENX%)
           JJ& = JJ& + 1
           Call GRAREG("HISTOVEN", Z$, JJ&)
           MsgBox JJ& & "   " & NCLIE% & "   " & VENX%
         End If
       End If
     Next UJ&
     Call CIERRARCH(ABRIX$)
   Next UI%
   Call CIERRARCH("HISTOVEN")
   '
End Sub

Sub LICOMIS()
    '
    Screen.MousePointer = 11
    '
    YALIQUI% = 0
    Dim CVX%(128), COVEN(128), COCOB(128)
    Dim PORCO(256), TOCOMI#(256), TOCOB#(256)
    Dim REGICLI&(32767), VENDECLI%(32767)
    Dim SUPED#(256), COMPED#(256), SUCOB#(256), COMCOB#(256)
    '
    Dim CMES$(13), FECOMIXI%(13)
    CMES$(1) = "NOV00": FECOMIXI%(1) = FECHANUM("01/11/00")
    CMES$(2) = "DIC00": FECOMIXI%(2) = FECHANUM("01/12/00")
    CMES$(3) = "ENE01": FECOMIXI%(3) = FECHANUM("01/01/01")
    CMES$(4) = "FEB01": FECOMIXI%(4) = FECHANUM("01/02/01")
    CMES$(5) = "MAR01": FECOMIXI%(5) = FECHANUM("01/03/01")
    CMES$(6) = "ABR01": FECOMIXI%(6) = FECHANUM("01/04/01")
    CMES$(7) = "MAY01": FECOMIXI%(7) = FECHANUM("01/05/01")
    CMES$(8) = "JUN01": FECOMIXI%(8) = FECHANUM("01/06/01")
    CMES$(9) = "JUL01": FECOMIXI%(9) = FECHANUM("01/07/01")
    CMES$(10) = "AGO01": FECOMIXI%(10) = FECHANUM("01/08/01")
    CMES$(11) = "SET01": FECOMIXI%(11) = FECHANUM("01/09/01")
    CMES$(12) = "OCT01": FECOMIXI%(12) = FECHANUM("01/10/01")
    CMES$(13) = "NOV01": FECOMIXI%(13) = FECHANUM("01/11/01")
    '
    If Option3.Value = True Then MODOCAL% = 2
    If Option4.Value = True Then MODOCAL% = 3
    If Option5.Value = True Then MODOCAL% = 4
    '
    DES% = FECHANUM("01/10/00")
    HAS% = FECHANUM("30/11/01")
    CORRESP$ = "01/10/00 - 30/11/01"
    '
    NCLIA% = 0
    JJ& = 0
    ' INDEXA CADA UNO DE LOS ARCHIVOS MENSUALES
    FIN& = 13
    For UJ& = 1 To 13
      Call TRACE(20, UJ&, FIN&)
      UKI% = UJ&
      ABRIX$ = "LCS" + CMES$(UKI%)
      Call COPYSTRU("LICOMIS", ABRIX$)
      X$ = REGLEIDO$(ABRIX$, 1)
      Call GRAREG(ABRIX$, X$, 1)
      Call CIERRARCH(ABRIX$)
    Next UJ&
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
50  'Call ARMAMES
    '
    Call HEADERS("LICOMIX", "")
    Call HEADERS("LICOMIX", "Corresponde a: " + CORRESP$)
    '
    Call COPYSTRU("DEUDOR1", "DEUVIA1")
    Call COPYSTRU("DEUDOR2", "DEUVIA2")
    ABRIX$ = "DEUVIA1"
    X$ = REGLEIDO$(ABRIX$, 1)
    Call GRAREG(ABRIX$, X$, 1)
    Call CIERRARCH(ABRIX$)
    Call COPYSTRU("DEUDOR1", "DEUVIB1")
    Call COPYSTRU("DEUDOR2", "DEUVIB2")
    ABRIX$ = "DEUVIB1"
    X$ = REGLEIDO$(ABRIX$, 1)
    Call GRAREG(ABRIX$, X$, 1)
    Call CIERRARCH(ABRIX$)
    Call COPYSTRU("DEUDOR1", "DEUVIC1")
    Call COPYSTRU("DEUDOR2", "DEUVIC2")
    ABRIX$ = "DEUVIC1"
    X$ = REGLEIDO$(ABRIX$, 1)
    Call GRAREG(ABRIX$, X$, 1)
    Call CIERRARCH(ABRIX$)
    '
    'FIN& = ULTREG&("DEUVIA1")
    'URE2& = ULTREG&("DEUVIA2")
    'For U& = 1 To FIN&
    '  Call TRACE(20, U&, FIN&)
    '  X$ = REGLEIDO$("DEUDOR1", U&)
    '  If U& <= URE2& Then
    '    X$ = X$ + REGLEIDO$("DEUDOR2", U&)
    '    NC% = CVI(Left$(X$, 2))
    '    If NC% > 0 Then
    '      If NC% <= 32767 Then
    '        If REGICLI&(NC%) < 1 Then
    '          REGICLI&(NC%) = U&
    '          VENDECLI%(NC%) = Asc(Mid$(X$, 152, 1))
    '        End If
    '      End If
    '    End If
    '  End If
    'Next U&
    '
    MODO% = 0
    If Option3.Value = True Then MODO% = 1
    If Option4.Value = True Then MODO% = 2
    If Option5.Value = True Then MODO% = 3
    '
    J& = 0: JLIBRO& = 0
    On MODO% GoTo 100, 200, 300
    GoTo 999
    '
100 Call HEADERS("LICOMIX", "Por: Pedidos")
    '
    FIN& = ULTREG&("PEDDETA")
    '
    INIPE& = 0: NCA% = 0
    For I& = 1 To FIN&
        Call TRACE(20, I& - INIPE&, FIN& - INIPE&)
        X$ = REGLEIDO$("PEDDETA", I&)
        FF% = CVI(Mid$(X$, 5, 2))
        If FF% >= DES% Then
            If FF% <= HAS% Then
                NRO& = CVS(Left$(X$, 4))
                If NRO& > 0 Then
                    NC% = CVI(Mid$(X$, 7, 2))
                    If NC% > 0 Then
                       VENDE% = VENDECLI%(NC%)
                       If NC% = NCA% Then GoTo 10
NCA% = NC%
For UKI% = 1 To 13
If FECOMIXI%(UKI%) >= FF% Then GoTo 5
Next UKI%
GoTo 8
5 ABRIX$ = "LCS" + CMES$(UKI%)
RFF$ = RECFILT$(ABRIX$, 3, MKI$(NC%))
If Len(RFF$) >= 4 Then
  REGE& = CVS(Left$(RFF$, 4))
  Y$ = REGLEIDO$(ABRIX$, REGE&)
  VENDE% = Asc(Left$(Y$, 1))
  VENDECLI%(NC%) = VENDE%
  Call CIERRARCH(ABRIX$)
  GoTo 10
End If
Call CIERRARCH(ABRIX$)
'
8  If VENDE% < 1 Then ' BUSCAR EN BACKUPS MAS NUEVOS DE DEUDOR1 / 2
     ABRIX$ = "DEUVIA"
     If FF% >= FECHANUM("02/10/00") Then ABRIX$ = "DEUVIB"
     If FF% >= FECHANUM("19/06/01") Then ABRIX$ = "DEUVIC"
     RFF$ = RECFILT$(ABRIX$ + "1", 1, MKI$(NC%))
     If Len(RFF$) >= 4 Then
       REGE& = CVS(Left$(RFF$, 4))
       UX$ = REGLEIDO$(ABRIX$ + "1", REGE&) + REGLEIDO$(ABRIX$ + "2", REGE&)
       VENDE% = Asc(Mid$(UX$, 152, 1))
       VENDECLI%(NC%) = VENDE%
     End If
     Call CIERRARCH(ABRIX$ + "1")
     Call CIERRARCH(ABRIX$ + "2")
   End If
   '
10 If VENDE% < 1 Then VENDE% = 99
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
    Next I&
    '
    GoTo 300                 ' PARA INCLUIR COBRANZA
    GoSub 610
    Call FINAL("*SORLICO")
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
                       If NC% = NCA% Then GoTo 20
NCA% = NC%
For UKI% = 1 To 13
If FECOMIXI%(UKI%) >= FF% Then GoTo 15
Next UKI%
GoTo 18
15 ABRIX$ = "LCS" + CMES$(UKI%)
RFF$ = RECFILT$(ABRIX$, 3, MKI$(NC%))
If Len(RFF$) >= 4 Then
  REGE& = CVS(Left$(RFF$, 4))
  Y$ = REGLEIDO$(ABRIX$, REGE&)
  VENDE% = Asc(Left$(Y$, 1))
  VENDECLI%(NC%) = VENDE%
  Call CIERRARCH(ABRIX$)
  GoTo 20
End If
Call CIERRARCH(ABRIX$)
'
18  If VENDE% < 1 Then ' BUSCAR EN BACKUPS MAS NUEVOS DE DEUDOR1 / 2
     ABRIX$ = "DEUVIA"
     If FF% >= FECHANUM("02/10/00") Then ABRIX$ = "DEUVIB"
     If FF% >= FECHANUM("19/06/01") Then ABRIX$ = "DEUVIC"
     RFF$ = RECFILT$(ABRIX$ + "1", 1, MKI$(NC%))
     If Len(RFF$) >= 4 Then
       REGE& = CVS(Left$(RFF$, 4))
       UX$ = REGLEIDO$(ABRIX$ + "1", REGE&) + REGLEIDO$(ABRIX$ + "2", REGE&)
       VENDE% = Asc(Mid$(UX$, 152, 1))
       VENDECLI%(NC%) = VENDE%
     End If
     Call CIERRARCH(ABRIX$ + "1")
     Call CIERRARCH(ABRIX$ + "2")
   End If
   '
20 If VENDE% < 1 Then VENDE% = 99
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
    Call FINAL("*SORLICO")
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
                       If NC% = NCA% Then GoTo 30
NCA% = NC%
For UKI% = 1 To 13
If FECOMIXI%(UKI%) >= FF% Then GoTo 25
Next UKI%
GoTo 28
25 ABRIX$ = "LCS" + CMES$(UKI%)
RFF$ = RECFILT$(ABRIX$, 3, MKI$(NC%))
If Len(RFF$) >= 4 Then
  REGE& = CVS(Left$(RFF$, 4))
  Y$ = REGLEIDO$(ABRIX$, REGE&)
  VENDE% = Asc(Left$(Y$, 1))
  VENDECLI%(NC%) = VENDE%
  Call CIERRARCH(ABRIX$)
  GoTo 30
End If
Call CIERRARCH(ABRIX$)
'
28  If VENDE% < 1 Then ' BUSCAR EN BACKUPS MAS NUEVOS DE DEUDOR1 / 2
     ABRIX$ = "DEUVIA"
     If FF% >= FECHANUM("02/10/00") Then ABRIX$ = "DEUVIB"
     If FF% >= FECHANUM("19/06/01") Then ABRIX$ = "DEUVIC"
     RFF$ = RECFILT$(ABRIX$ + "1", 1, MKI$(NC%))
     If Len(RFF$) >= 4 Then
       REGE& = CVS(Left$(RFF$, 4))
       UX$ = REGLEIDO$(ABRIX$ + "1", REGE&) + REGLEIDO$(ABRIX$ + "2", REGE&)
       VENDE% = Asc(Mid$(UX$, 152, 1))
       VENDECLI%(NC%) = VENDE%
     End If
     Call CIERRARCH(ABRIX$ + "1")
     Call CIERRARCH(ABRIX$ + "2")
   End If
   '
30 If VENDE% < 1 Then VENDE% = 99
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
320                       ICOM# = (Int(ITOT# * PORCOM + 0.5)) / 100
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
    Call FINAL("*SORLICO")
    YALIQUI% = 1
    GoTo 999
    '
    '
999 If YALIQUI% = 1 Then
        FORLICO.List1.Clear
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
                FORLICO.List1.AddItem TXT$
            End If
        Next VENDE%
        FORLICO.Check4.Value = 1
        '
        If FORLICO.List1.ListCount > 0 Then
            FORLICO.List1.ListIndex = 0
        End If
        FORLICO.List1.Refresh
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

Sub PRILICOMIS()
    '
    MODOCAL% = 0
    If FORLICO.Option2.Value = True Then MODOCAL% = 1
    If FORLICO.Option3.Value = True Then MODOCAL% = 2: MODOCA$ = "Por Pedidos Netos de Descuento"
    If FORLICO.Option4.Value = True Then MODOCAL% = 3: MODOCA$ = "Por Facturacion"
    If FORLICO.Option5.Value = True Then MODOCAL% = 4: MODOCA$ = "Por Cobranza Neta de IVA"
    If MODOCAL% < 1 Then
        Exit Sub
    End If
    '
    CORRESP$ = "01/10/00 - 30/11/01"
    Call SELECHO("VENDEDOR")
    VENDE% = Val(VALACT1$("VENDEDOR"))
    If VENDE% < 1 Then Exit Sub
    
            J& = 0
            Call HEADERS("LICOMIX", "")
            
           NOVEN$ = ECOARCH$("VENDEDOR", Chr$(VENDE%))
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
            Call CONECRUN("*SORLIQX", "Listado de Comisiones Liquidadas")
            Exit Sub
    '
End Sub
'

