VERSION 5.00
Begin VB.Form DISTRIMAT 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asignación de Pedidos - Distribución de Materiales"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   8610
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   8610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   7600
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   23
      Top             =   4410
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
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   210
      TabIndex        =   21
      Top             =   3105
      Width           =   7170
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Stock Neto"
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
      Left            =   4730
      TabIndex        =   16
      Top             =   1365
      Width           =   2640
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   1260
         TabIndex        =   20
         Top             =   315
         Width           =   1170
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "St.Neto:"
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
         Left            =   315
         TabIndex        =   19
         Top             =   420
         Width           =   855
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   1260
         TabIndex        =   18
         Top             =   735
         Width           =   1170
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Disponible:"
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
         Left            =   105
         TabIndex        =   17
         Top             =   840
         Width           =   1065
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Pedidos"
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
      Left            =   2040
      TabIndex        =   11
      Top             =   1365
      Width           =   2430
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Asignado:"
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
         Left            =   105
         TabIndex        =   15
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   1050
         TabIndex        =   14
         Top             =   735
         Width           =   1170
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
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
         Left            =   105
         TabIndex        =   13
         Top             =   420
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   1050
         TabIndex        =   12
         Top             =   315
         Width           =   1170
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Stock"
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
      Left            =   210
      TabIndex        =   9
      Top             =   1365
      Width           =   1590
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   210
         TabIndex        =   10
         Top             =   735
         Width           =   1170
      End
   End
   Begin VB.TextBox Text1 
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
      Left            =   1080
      TabIndex        =   0
      Top             =   315
      Width           =   2115
   End
   Begin VB.TextBox Text2 
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
      Left            =   6720
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   315
      Width           =   645
   End
   Begin VB.TextBox Text3 
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
      Left            =   210
      MultiLine       =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   840
      Width           =   7170
   End
   Begin VB.CommandButton Command2 
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
      Height          =   360
      Left            =   3150
      TabIndex        =   1
      Top             =   315
      Width           =   175
   End
   Begin VB.CommandButton Command5 
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
      Left            =   7600
      Picture         =   "DISTRIMAT.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
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
      Left            =   7600
      Picture         =   "DISTRIMAT.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Terminar - Salir de Consulta de Stocks"
      Top             =   315
      Width           =   855
   End
   Begin VB.CommandButton COMMAND3 
      Caption         =   "Kardex"
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
      Left            =   7600
      Picture         =   "DISTRIMAT.frx":0454
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Ficha Tipo Kardex de Movimientos Históricos"
      Top             =   2205
      Width           =   855
   End
   Begin VB.Line Line2 
      X1              =   7380
      X2              =   7380
      Y1              =   2835
      Y2              =   5340
   End
   Begin VB.Line Line1 
      X1              =   210
      X2              =   8700
      Y1              =   5325
      Y2              =   5325
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   7035
      Picture         =   "DISTRIMAT.frx":075E
      Top             =   4810
      Width           =   2010
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Nro    Cliente / Referencia                               Pedido        Asignado       Neto         "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   210
      TabIndex        =   22
      Top             =   2835
      Width           =   7170
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
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
      Left            =   0
      TabIndex        =   8
      Top             =   420
      Width           =   960
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "U-M"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5760
      TabIndex        =   7
      Top             =   420
      Width           =   855
   End
End
Attribute VB_Name = "DISTRIMAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    DISTRIMAT.Hide
End Sub

Private Sub Command2_Click()
    '
    Call SELECHO("CODISTRI")
    CODD$ = VALACT1$("CODISTRI")
    If CODD$ = "" Then
        Exit Sub
    End If
    COD$ = CODD$
    Text1.TEXT = COD$
    CI% = CODINT%(COD$)
    Text3.TEXT = DESCRIT$(CI%)
    Text2.TEXT = UNIMED$(CI%)
    Call MUEDISTRI(CI%)                ' recalcular existencias
    '
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    Call GENKARDEX(COD$)
    Command3.Enabled = True
End Sub

Private Sub Command5_Click()
    Call MENSERR(24, "No se Encuentra Ayuda Específica")
End Sub

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    LU$ = LUDAT$
    '
End Sub
'

Private Sub Text1_DblClick()
    '
    Call SELECHO("CODISTRI")
    CODD$ = VALACT1$("CODISTRI")
    If CODD$ = "" Then
        Exit Sub
    End If
    COD$ = CODD$
    Text1.TEXT = COD$
    CI% = CODINT%(COD$)
    Text3.TEXT = DESCRIT$(CI%)
    Text2.TEXT = UNIMED$(CI%)
    Call MUEDISTRI(CI%)                ' recalcular existencias
    '
End Sub
'

Private Sub Text1_KeyPress(KeyAscii As Integer)
    '
    If KeyAscii = 13 Then
        Text2.TEXT = ""
        Text3.TEXT = ""
        Label4.Caption = ""
        Label5.Caption = ""
        Label7.Caption = ""
        Label10.Caption = ""
        Label14.Caption = ""
        List2.Clear
        If CODINT%(Text1.TEXT) > 0 Then
            CODD$ = AJUSTI$(Text1.TEXT, 16)
            If ECOARCH$("CODISTRI", CODD$) <> "" Then
                COD$ = TRIM$(Text1.TEXT)
                CI% = CODINT%(COD$)
                Text2.TEXT = UNIMED$(CI%)
                Text3.TEXT = DESCRIT$(CI%)
                Call MUEDISTRI(CI%)                ' recalcular existencias
            End If
        End If
    End If
    '
    
End Sub

Private Sub Text1_LostFocus()
    CODD$ = AJUSTI$(Text1.TEXT, 16)
    If ECOARCH$("CODISTRI", CODD$) <> "" Then
        COD$ = TRIM$(Text1.TEXT)
        CI% = CODINT%(COD$)
        Text2.TEXT = UNIMED$(CI%)
        Text3.TEXT = DESCRIT$(CI%)
    End If
End Sub

Private Sub Text2_GotFocus()
    Text1.SetFocus
End Sub

Private Sub Text3_Click()
    Text1.SetFocus
End Sub

Sub MUEDISTRI(CI%)
    '
    If CI% < 1 Or CI% > NUMAS% Then
        Exit Sub
    End If
    '
    Screen.MousePointer = 11
    '
    Static NUSU%, COSU%(10), SUCU$(10)
    If NUSU% = 0 Then
        For U& = 1 To ULTREG&("SUCURSAL")
           If U& <= 10 Then
              SS$ = REGLEIDO$("SUCURSAL", U&)
              If Asc(Left$(SS$, 1)) > 32 Then
                 NUSU% = U&
                 COSU%(U&) = Asc(Mid$(SS$, 64, 1))
                 SUCU$(U&) = Left$(SS$, 20)
              End If
           End If
        Next U&
     End If
     '
     STOTOT = 0
     '
     REG& = CI% + 1
     If REG& > 0 Then
        If REG& <= NUMAS% Then
           SALCA$ = REGLEIDO$("SALSTOCK", REG&)
           For U& = 1 To NUSU%
               XXZ$ = Mid$(SALCA$, 4 * U& - 3, 4)
               XXX$ = SUCU$(U&) + CSTRING$(XXZ$, 4, 12, 1, 2)
               STOTOT = STOTOT + CVS(XXZ$)
           Next U&
           DISTRIMAT.Label4.Caption = TRIM$(CSTRING$(MKS$(STOTOT), 4, 12, 1, 2))
       End If
     End If
     '
10   DISTRIMAT.List2.Clear
     HOII% = HOY(HO$)
     SALSTO = STOTOT
     '
     RGXX$ = RECFILT$("PEDDETP", 12, MKI$(CI%))
     PEDI = 0: ASII = 0: TOASI = 0
     For I& = 1 To Len(RGXX$) Step 4
        REG& = CVS(Mid$(RGXX$, I&, 4))
        X$ = REGLEIDO$("PEDDETP", REG&)
        If Asc(Mid$(X$, 27, 1)) <= 1 Then
            CIP% = CVI(Mid$(X$, 33, 2))
            If CIP% <> CI% Then        ' fuerza reindexacion
                X$ = REGLEIDO$("PEDDETP", 1)
                Call GRAREG("PEDDETP", X$, 1)
                Call CIERRARCH("PEDDETP")
                GoTo 10
            End If
            '
            NROPED& = CVS(Left$(X$, 4))
            NCI% = CVI(Mid$(X$, 7, 2))
            CASIG = 0
            If Mid$(X$, 75, 4) <> Space$(4) Then
                CASIG = CVS(Mid$(X$, 75, 4))
            End If
            CAPEN = CVS(Mid$(X$, 79, 4)) - CVS(Mid$(X$, 95, 4))
            If CAPEN < 0 Then CAPEN = 0
            '
            If ASII + CASIG > STOTOT Then
                CASIG = STOTOT - ASII
            End If
            '
            If CAPEN > 0 Or CASIG > 0 Then
                PEDI = PEDI + CAPEN
                ASII = ASII + CASIG
                If ASII > PEDI Then ASII = PEDI
            End If
            '
            SALSTO = SALSTO - CAPEN
            '
        End If
     Next I&
     If ASII > STOTOT Then
        ASII = STOTOT
     End If
     '
     DISTRIMAT.Label5.Caption = CSTRING$(MKS$(PEDI), 4, 12, 2, 2)
     DISTRIMAT.Label7.Caption = CSTRING$(MKS$(ASII), 4, 12, 2, 2)
     DISTRIMAT.Label14.Caption = CSTRING$(MKS$(STOTOT - PEDI), 4, 12, 2, 2)
     DISTRIMAT.Label10.Caption = CSTRING$(MKS$(STOTOT - ASII), 4, 12, 2, 2)
     '
     Call BLANARCH("!DISTRIMA")
     For I& = 1 To ULTREG&("ODESEMI")
        X$ = REGLEIDO$("ODESEMI", I&)
        If CVI(Mid$(X$, 11, 2)) = CI% Then
            '
            NROPED& = CVS(Left$(X$, 4))
            NC% = CVI(Mid$(X$, 7, 2))
            CANPE = CVS(Mid$(X$, 13, 4))
            CANAS = CVS(Mid$(X$, 17, 4))
            '
            Z$ = REGBLAN$("DISTRIMA")
            Call REPLA(Z$, MKS$(NROPED&), 1, 4)
            Call REPLA(Z$, RASOCLI$(NC%), 5, 30)
            Call REPLA(Z$, MKS$(CANPE), 35, 4)
            Call REPLA(Z$, MKS$(CANAS), 39, 4)
            Call REPLA(Z$, MKS$(I&), 43, 4)
            Call REGAPP("!DISTRIMA", Z$)
            '
        End If
     Next I&
     '
     Screen.MousePointer = 1
     '
     ATRI$ = "/NC"
     ATRI$ = ATRI$ + "/TITUPAN=Asignacion de Pedidos - Redistribucion"
     ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(DISTRIMAT.Top + DISTRIMAT.Label15.Top))
     ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(DISTRIMAT.Left + DISTRIMAT.Label15.Left))
     '
20   VTB% = VENTABU%("DISTRIMA" + ATRI$)
     If VTB% >= 0 Then
        SUMASI = 0
        For I& = 1 To ULTREG&("!DISTRIMA")
            X$ = REGLEIDO$("!DISTRIMA", I&)
            If CVS(Mid$(X$, 39, 4)) > CVS(Mid$(X$, 35, 4)) Then
                Call MENSERR(24, "La Cantidad Asignada\Supera la Pedida.")
                GoTo 20
            End If
            SUMASI = SUMASI + CVS(Mid$(X$, 39, 4))
        Next I&
        '
        If SUMASI > Val(DISTRIMAT.Label10.Caption) Then
            Call MENSERR(24, "Las Cantidades Distribuidas\Superan el Stock Disponible")
            GoTo 20
        End If
        '
        For I& = 1 To ULTREG&("!DISTRIMA")
            X$ = REGLEIDO$("!DISTRIMA", I&)
            CANAS = CVS(Mid$(X$, 39, 4))
            REGODES& = CVS(Mid$(X$, 43, 4))
            Z$ = REGLEIDO$("ODESEMI", REGODES&)
            Call REPLA(Z$, MKS$(CANAS), 17, 4)
            Call GRAREG("ODESEMI", Z$, REGODES&)
        Next I&
        '
        Call CIERRARCH("ODESEMI")
        '
     End If
     '
     Call CIERRARCH("!DISTRIMA")
     Call BLANARCH("!DISTRIMA")
     '
End Sub
'

