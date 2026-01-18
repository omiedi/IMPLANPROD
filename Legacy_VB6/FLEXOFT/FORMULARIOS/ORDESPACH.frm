VERSION 5.00
Begin VB.Form ORDESPACH 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asignación de Pedidos - Ordenes de Despacho"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5610
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
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
      Left            =   4560
      Picture         =   "ORDESPACH.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1320
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFF80&
      Caption         =   "Actualización"
      Height          =   3135
      Left            =   240
      TabIndex        =   9
      Top             =   240
      Width           =   4110
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1680
         ScaleHeight     =   585
         ScaleWidth      =   2175
         TabIndex        =   25
         Top             =   2280
         Width           =   2235
         Begin VB.CommandButton Command8 
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
            Picture         =   "ORDESPACH.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Anulación de O./Despacho"
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
            Left            =   720
            TabIndex        =   27
            Top             =   120
            Width           =   1455
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1320
         ScaleHeight     =   585
         ScaleWidth      =   2175
         TabIndex        =   22
         Top             =   1800
         Width           =   2235
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
            Left            =   -20
            Picture         =   "ORDESPACH.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Repetir Impre- sion de O/D"
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
            TabIndex        =   24
            Top             =   0
            Width           =   1455
         End
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   960
         ScaleHeight     =   585
         ScaleWidth      =   2175
         TabIndex        =   10
         Top             =   1320
         Width           =   2235
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
            Picture         =   "ORDESPACH.frx":0C7E
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Emitir Ordenes de Despacho"
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
            Left            =   720
            TabIndex        =   12
            Top             =   0
            Width           =   1455
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   600
         ScaleHeight     =   585
         ScaleWidth      =   2175
         TabIndex        =   13
         Top             =   840
         Width           =   2235
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
            Left            =   0
            Picture         =   "ORDESPACH.frx":0F88
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Re-Distribucion de Materiales"
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
            Left            =   720
            TabIndex        =   15
            Top             =   0
            Width           =   1455
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   300
         ScaleHeight     =   585
         ScaleWidth      =   2175
         TabIndex        =   16
         Top             =   360
         Width           =   2235
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
            Left            =   -20
            Picture         =   "ORDESPACH.frx":10D2
            Style           =   1  'Graphical
            TabIndex        =   0
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Revisar Situa- ción de Stocks"
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
            TabIndex        =   17
            Top             =   0
            Width           =   1590
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Consultas - Listados"
      Height          =   3210
      Left            =   240
      TabIndex        =   2
      Top             =   3600
      Width           =   5160
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFF80&
         Caption         =   "O/Despacho en Curso por Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   525
         TabIndex        =   28
         Top             =   840
         Width           =   4530
      End
      Begin VB.OptionButton Option7 
         BackColor       =   &H00FFFF80&
         Caption         =   "Pedidos sin Entrega - Falta de Stock"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   525
         TabIndex        =   21
         Top             =   1470
         Width           =   4275
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   240
         ScaleHeight     =   75
         ScaleWidth      =   4695
         TabIndex        =   19
         Top             =   2835
         Width           =   4750
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   20
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   4095
         Picture         =   "ORDESPACH.frx":13DC
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2100
         Width           =   855
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00FFFF80&
         Caption         =   "Ordenes de Despacho en Curso por Número"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   525
         TabIndex        =   7
         Top             =   525
         Width           =   4530
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFF80&
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
         TabIndex        =   4
         Top             =   2020
         Width           =   3690
         Begin VB.OptionButton Option5 
            BackColor       =   &H00FFFF80&
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
            Left            =   315
            TabIndex        =   6
            Top             =   250
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.OptionButton Option4 
            BackColor       =   &H00FFFF80&
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
            Left            =   2100
            TabIndex        =   5
            Top             =   250
            Width           =   1275
         End
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFF80&
         Caption         =   "Material Disponible para Asignar (O/D's a Emitir)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   520
         TabIndex        =   3
         Top             =   1155
         Value           =   -1  'True
         Width           =   4500
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
      Height          =   855
      Left            =   4540
      Picture         =   "ORDESPACH.frx":16E6
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   555
      Left            =   4320
      Picture         =   "ORDESPACH.frx":1830
      Stretch         =   -1  'True
      Top             =   2800
      Width           =   5250
   End
End
Attribute VB_Name = "ORDESPACH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command2_Click()
    OPCI% = 1
    If Option2.Value = True Then
       Call CONODESCLI
       Exit Sub
    End If
    '
    If Option6.Value = True Then OPCI% = 4
    If Option7.Value = True Then OPCI% = 5
    If OPCI% = 1 Then
        If Option5.Value = True Then
             Call CONECRUN("*CODESEM", "")
           Else
             Call CONECRUN("*ODESEMI", "")
        End If
      ElseIf OPCI% = 4 Then
        Screen.MousePointer = 11
        Call COPYSTRU("ORDESPA", "ORDESPE")
        '
        FIN& = ULTREG&("ORDESPA")
        JJ& = 0
        For U& = 1 To FIN&
          Call TRACE(20, U&, FIN&)
          X$ = REGLEIDO$("ORDESPA", U&)
          If Asc(Mid$(X$, 33, 1)) <> 9 Then ' no esta anulado
            NODES& = CVS(Left$(X$, 4))
            If NODES& > 0 Then
              NPED& = CVS(Mid$(X$, 7, 4))
              If NPED& > 0 Then
                CAPEN = CVS(Mid$(X$, 79, 4)) - CVS(Mid$(X$, 83, 4))
                If CAPEN >= 0.05 Then
                  CI% = CVI(Mid$(X$, 37, 2))
                  If CI% > 0 Then
                    If CI% <= NUMAS% Then
                      JJ& = JJ& + 1
                      Call GRAREG("ORDESPE", X$, JJ&)
                    End If
                  End If
                End If
              End If
            End If
          End If
        Next U&
        Call SETULTREG("ORDESPE", JJ&)
        Screen.MousePointer = 1
        '
        If JJ& < 1 Then
          Call MENSERR(24, "No Hay Ordenes de Despacho en Curso")
          Exit Sub
        End If
        '
        If Option5.Value = True Then
             Call CONECRUN("*COODESP", "")
           Else
             Call CONECRUN("*LIODESP", "")
        End If
        '
      ElseIf OPCI% = 5 Then
        Call CALPESTOCE
        If Option5.Value = True Then
             Call CONECRUN("*SOSTOCC", "")
           Else
             Call CONECRUN("*SOSTOCE", "")
        End If
      Else
        Hide
    End If
    '
End Sub

Private Sub Command3_Click()
    COMMAND3.Enabled = False
    Call CALASIAUT
    COMMAND3.Enabled = True
End Sub

Private Sub Command3_GotFocus()
   If REVIPEPE% = 1 Then
       Refresh
       Call CALASIAUT
   End If
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    DISTRIMAT.Show 1
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    Call EMIODES
    Command5.Enabled = True
End Sub

Private Sub Command6_Click()
    Call HELPONLINE("RESERVA")
End Sub

Private Sub Command7_Click()
    Command7.Enabled = False
    Screen.MousePointer = 11
    JJ& = 0
    For U& = 1 To ULTREG&("ORDESPA")
        X$ = REGLEIDO$("ORDESPA", U&)
        NODES& = CVS(Left$(X$, 4))
        If NODES& > 0 Then
           If NODES& <> NODESA& Then
              NODESA& = NODES&
              z$ = REGBLAN$("INDIODES")
              NC% = CVI(Mid$(X$, 13, 2))
              CLI$ = "": If NC% > 0 Then CLI$ = RASOCLI$(NC%)
              FEX = CVI(Mid$(X$, 5, 2))
              Z1$ = FECHATEX$(FEX) + "  " + CLI$
              '
              Call REPLA(z$, MKS$(NODES&), 1, 4)
              Call REPLA(z$, Z1$, 5, 40)
              JJ& = JJ& + 1
              Call GRAREG("INDIODES", z$, JJ&)
           End If
        End If
    Next U&
    '
    Call SETULTREG("INDIODES", JJ&)
    Call CIERRARCH("INDIODES")
    Call FRESHECHO("INDIODES")
    '
10  Call SELECHO("INDIODES")
    NODES& = Val(VALACT1$("INDIODES"))
    If NODES& > 0 Then
        Call PRIODESPA(NODES&)
        GoTo 10
    End If
    '
    Command7.Enabled = True
    '
End Sub

Private Sub Command8_Click()
   '
   Command8.Enabled = False
5  Screen.MousePointer = 11
   FIN& = ULTREG&("ORDESPA")
   JJ& = 0
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     X$ = REGLEIDO$("ORDESPA", U&)
     If Asc(Mid$(X$, 33, 1)) <> 9 Then ' no esta anulado
       NODES& = CVS(Left$(X$, 4))
       If NODES& > 0 Then
         NPED& = CVS(Mid$(X$, 7, 4))
         If NPED& > 0 Then
           CAPEN = CVS(Mid$(X$, 79, 4)) - CVS(Mid$(X$, 83, 4))
           If CAPEN >= 0.05 Then
             CI% = CVI(Mid$(X$, 37, 2))
             If CI% > 0 Then
               If CI% <= NUMAS% Then
                 If NODES& <> NODESA& Then
                   NODESA& = NODES&
                   z$ = REGBLAN$("INDIODES")
                   NC% = CVI(Mid$(X$, 13, 2))
                   CLI$ = "": If NC% > 0 Then CLI$ = RASOCLI$(NC%)
                   FEX = CVI(Mid$(X$, 5, 2))
                   Z1$ = FECHATEX$(FEX) + "  " + CLI$
                   '
                   Call REPLA(z$, MKS$(NODES&), 1, 4)
                   Call REPLA(z$, Z1$, 5, 40)
                   JJ& = JJ& + 1
                   Call GRAREG("INDIODES", z$, JJ&)
                 End If
               End If
             End If
           End If
         End If
       End If
     End If
   Next U&
   '
   Call SETULTREG("INDIODES", JJ&)
   Call CIERRARCH("INDIODES")
   Call FRESHECHO("INDIODES")
   '
10 Screen.MousePointer = 1
   Call SELECHO("INDIODES")
   NODES& = Val(VALACT1$("INDIODES"))
   If NODES& > 0 Then
     Call ANULODESPA(NODES&)
     GoTo 5
   End If
   '
   Command8.Enabled = True
   '
End Sub

Private Sub Form_Load()
   Option1.Value = True
End Sub

Sub CALASIAUT()          ' revision de situacion de stocks
    '                      para asignacion de pedidos
    '
    Screen.MousePointer = 11
    Dim PUNPED$(16384)
    Dim STOLE%(16384)
    Dim STTOT(16384)
    Dim STASI(16384)
    '
    REVIPEPE% = 0
    NPEDA& = NPED&
    NODEA& = NODE&
    '
    HOI% = HOY(HOS$)
    DIANTIDES% = Val(CONTROL$("", "ANTIDESP"))
    If DIANTIDES% < 1 Then DIANTIDES% = 3
    '
    FIN& = ULTREG&("PEDDETA")
    '
    ' revisar y depurar pedidos pendientes
    '
    V1& = 0: YABLOPEP% = 0
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("PEDDETA", U&)
      If Asc(Mid$(X$, 27, 1)) <> 9 Then ' no esta anulado
        CAN = CVS(Mid$(X$, 79, 4))
        If CAN >= 0.05 Then
          CADESPACH = CVS(Mid$(X$, 95, 4))
          CAPENDENT = CAN - CADESPACH
          If CAPENDENT > 0.5 Then
            NCLIE% = CVI(Mid$(X$, 7, 2))
            NPED& = CVS(Left$(X$, 4))
            If NPED& > 0 Then
              CI% = CVI(Mid$(X$, 33, 2))
              REGENCA& = CVS(Mid$(X$, 109, 4))
              '
              z$ = REGLEIDO$("PEDENCA", REGENCA&)
              NPEDA& = CVS(Left$(z$, 4))
              NCLIA% = CVI(Mid$(z$, 55, 2))
              If NPEDA& <> NPED& Or NCLIA% <> NCLIE% Then
                Call MENSERR(24, "Punteros Incorrectos en Pedido " + TRIM$(Str$(NPED&)) + "/" + TRIM$(Str$(NPEDA&)) + "\Consulte a su Soporte Técnico")
                GoTo 90
              End If
              '
              If Asc(Mid$(z$, 267, 1)) > 1 Then
                Call REPLA(X$, Chr$(9), 27, 1)  ' anula detalle
                Call GRAREG("PEDDETA", X$, U&)
                GoTo 90
              End If
              '
              Y$ = X$
              Call REPLA(Y$, MKS$(0), 75, 4) ' pone en cero cantidad asignada
              Call REPLA(Y$, MKS$(U&), 109, 4)
              V1& = V1& + 1
              Call GRAREG("PEDDETP", Y$, V1&)
              '
              If CI% > 0 Then
                If CI% <= NUMAS% Then
                  PUNPED$(CI%) = PUNPED$(CI%) + MKS$(V1&)
                End If
              End If
              '
            End If
          End If
        End If
      End If
      '
90  Next U&
    Call SETULTREG("PEDDETP", V1&)
    Call CIERRARCH("PEDDETP")
    '
    ' depurar segun ordenes de despacho en curso
    '
    UREPEP& = ULTREG&("PEDDETP")
    FIN& = ULTREG&("ORDESPA")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("ORDESPA", U&)
      '
      If Asc(Mid$(X$, 33, 1)) <> 9 Then ' no esta anulado
        NODE& = CVS(Left$(X$, 4))
        NPED& = CVS(Mid$(X$, 7, 4))
        If NPED& > 0 Then
          CAPEN = CVS(Mid$(X$, 79, 4)) - CVS(Mid$(X$, 83, 4))
          If CAPEN >= 0.05 Then
            CI% = CVI(Mid$(X$, 37, 2))
            If CI% > 0 Then
              If CI% <= NUMAS% Then
                For II% = 1 To Len(PUNPED$(CI%)) Step 4
                  REGPEDP& = CVS(Mid$(PUNPED$(CI%), II%, 4))
                  If REGPEDP& > 0 Then
                    If REGPEDP& <= UREPEP& Then
                      Y$ = REGLEIDO$("PEDDETP", REGPEDP&)
                      NPEDA& = CVS(Left$(Y$, 4))
                      If NPEDA& <> NPED& Then
                        GoTo 95
                      End If
                      '
                      CAPENPE = CVS(Mid$(Y$, 79, 4)) - CVS(Mid$(Y$, 95, 4))
                      CASIG = CAPEN
                      If CASIG > CAPENPE Then
                        CASIG = CAPENPE
                      End If
                      '
                      CAPEN = CAPEN - CASIG
                      TOASI = CVS(Mid$(Y$, 75, 4)) + CASIG
                      Call REPLA(Y$, MKS$(TOASI), 75, 4)
                      Call GRAREG("PEDDETP", Y$, REGPEDP&)
                      '
                      REGPEDA& = CVS(Mid$(Y$, 109, 4))
                      Y$ = REGLEIDO$("PEDDETA", REGPEDA&)
                      Call REPLA(Y$, MKS$(TOASI), 75, 4)
                      Call GRAREG("PEDDETA", Y$, REGPEDA&)
                      '
                      If STOLE%(CI%) = 0 Then
                         STOLE%(CI%) = 1
                         STTOT(CI%) = STODEPOS(CI%, DEPOEXPE%)
                      End If
                      STTOT(CI%) = STTOT(CI%) - CASIG
                      GoTo 99
                      '
                    End If
                  End If
95              Next II%
                '
              End If
            End If
          End If
        End If
      End If
      '
99  Next U&
    '
    FIN& = ULTREG&("PEDDETP")
    JJ& = 0: KJ& = 0
    For U& = 1 To FIN&
        Call TRACE(20, U&, FIN&)
        X$ = REGLEIDO$("PEDDETP", U&)
        CAPENPE = CVS(Mid$(X$, 79, 4)) - CVS(Mid$(X$, 75, 4)) - CVS(Mid$(X$, 95, 4))
        If CAPENPE > 0.05 Then
          CI% = CVI(Mid$(X$, 33, 2))
          NPED& = CVS(Left$(X$, 4))
          FEPE% = CVI(Mid$(X$, 5, 2))
          NC% = CVI(Mid$(X$, 7, 2))
          FENTRESOL% = CVI(Mid$(X$, 25, 2))
          REGPEDA& = CVS(Mid$(X$, 109, 4))
          '
          FE1 = HOI%: FE2 = FENTRESOL%
          If DIENTRE%(FE1, FE2) <= DIANTIDES% Then
            If CI% > 0 Then
              If CI% <= NUMAS% Then
                '
                If STOLE%(CI%) = 0 Then
                   STOLE%(CI%) = 1
                   STTOT(CI%) = STODEPOS(CI%, DEPOEXPE%)
                End If
                '
                If STTOT(CI%) >= 0.05 Then
                  Y$ = REGBLAN$("ODESEMI")
                  Call REPLA(Y$, MKS$(NPED&), 1, 4)
                  Call REPLA(Y$, MKI$(FEPE%), 5, 2)
                  Call REPLA(Y$, MKI$(NC%), 7, 2)
                  Call REPLA(Y$, MKI$(FENTRESOL%), 9, 2)
                  Call REPLA(Y$, MKI$(CI%), 11, 2)
                  Call REPLA(Y$, MKS$(CAPENPE), 13, 4)
                  '
                  CANASI = CAPENPE
                  If CANASI > STTOT(CI%) - STASI(CI%) Then
                      CANASI = STTOT(CI%) - STASI(CI%)
                      CODD$ = TRIM$(CODEXT$(CI%))
                      For UVX& = 1 To KJ&
                         ZZ$ = REGLEIDO$("CODISTRI", UVX&)
                         If TRIM$(Left$(ZZ$, 16)) = CODD$ Then GoTo 100
                      Next UVX&
                      ZZ$ = REGBLAN$("CODISTRI")
                      KJ& = KJ& + 1
                      Call REPLA(ZZ$, CODD$, 1, 16)
                      Call REPLA(ZZ$, DESCRIT$(CI%), 17, 48)
                      Call GRAREG("CODISTRI", ZZ$, KJ&)
                  End If
                  '
100               Call REPLA(Y$, MKS$(CANASI), 17, 4)
                  Call REPLA(Y$, MKS$(REGPEDA&), 21, 4)
                  '
                  JJ& = JJ& + 1
                  Call GRAREG("ODESEMI", Y$, JJ&)
                  '
                  STASI(CI%) = STASI(CI%) + CANASI
                  '
                End If
              End If
            End If
          End If
        End If
    Next U&
    '
    Call SETULTREG("ODESEMI", JJ&)
    Call CIERRARCH("ODESEMI")
    Call SETULTREG("CODISTRI", KJ&)
    Call CIERRARCH("CODISTRI")
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Screen.MousePointer = 1
    '
    TTXX$ = "Proceso de Revisión Completo.\  \Puede Ahora Emitir Listado\de Pedidos a Despachar."
    If COMALTER%(TTXX$ + "\\Emitir Listado\Cancelar") = 1 Then
        Call CONECRUN("*ODESEMI", "")
    End If
    '
End Sub
'
Sub EMIODES()
    '
    Screen.MousePointer = 11
    FODESEMI.LIPEDES.Clear
    '
    NPEDA& = 0
    For U& = 1 To ULTREG&("ODESEMI")
      X$ = REGLEIDO$("ODESEMI", U&)
      NPED& = CVS(Left$(X$, 4))
      If NPED& > 0 Then
        If NPED& <> NPEDA& Then
          If CVS(Mid$(X$, 17, 4)) >= 0.05 Then
            NPEDA& = NPED&
            NC% = CVI(Mid$(X$, 7, 2))
            z$ = CSTRING$(MKS$(NPED&), 3, 7, 0, 2) + "  " + RASOCLI$(NC%)
            FODESEMI.LIPEDES.AddItem z$
          End If
        End If
      End If
    Next U&
    '
    Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    If FODESEMI.LIPEDES.ListCount < 1 Then
      Call COMUNI("No Hay Pedidos Pendientes\que Puedan Asignarse.")
      Exit Sub
    End If
    '
    FODESEMI.Show 1
    Call FRESHALL
    '
    Screen.MousePointer = 11
    For U& = 1 To FODESEMI.LIPEDES.ListCount
      If FODESEMI.LIPEDES.Selected(U& - 1) = True Then
        FODESEMI.LIPEDES.ListIndex = U& - 1
        CAPESEL% = CAPESEL% + 1
        TTXX$ = FODESEMI.LIPEDES.TEXT
        '
        NPED& = Val(Left$(TTXX$, 7))
20      Call GENODESPA(NPED&, NODES&)
        If NODES& > 0 Then
          Call PRIODESPA(NODES&)
        End If
        If NPED& < 0 Then
          NPED& = Abs(NPED&)
          GoTo 20
        End If
      End If
    Next U&
    '
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    '
    Screen.MousePointer = 1
    '
End Sub
'
Sub GENODESPA(NPED&, NODES&)
    '
    Dim REODES&(1024)
    KKK% = 0
    '
    Call BLOQARCH("ORDESPA")
    '
    FORIPRE$ = TRIM$(CONTROL$("", "FORDESPA"))
    If FORIPRE$ = "" Then
        FORIPRE$ = "ODES-SAB"
    End If
    CODFOR$ = TRIM$(UCase$(FORIPRE$))
    MAXIT% = Val(ATRIFORM$(CODFOR$, "MAX-ITEM"))
    If MAXIT% < 1 Or MAXIT% > 256 Then MAXIT% = 10
    '
    NODES& = 0
    For U& = 1 To ULTREG&("ODESEMI")
      X$ = REGLEIDO$("ODESEMI", U&)
      If CVS(Left$(X$, 4)) = NPED& Then
        CI% = CVI(Mid$(X$, 11, 2))
        Call VESARRAS(CI%)
        CANASI = CVS(Mid$(X$, 17, 4))
        STAXX = STODEPOS(CI%, DEPOEXPE%)
        '
        If STAXX < 0 Then STAXX = 0
        If CANASI > STAXX Then CANASI = STAXX
        '
        If CANASI >= 0.05 Then
          REGPEDA& = CVS(Mid$(X$, 21, 4))
          If REGPEDA& > 0 Then
            If REGPEDA& <= ULTREG&("PEDDETA") Then
              Y$ = REGLEIDO$("PEDDETA", REGPEDA&)
              If CVS(Left$(Y$, 4)) = NPED& Then
                If CVI(Mid$(Y$, 33, 2)) = CI% Then
                  KKK% = KKK% + 1
                  REODES&(KKK%) = U&
                  GoTo 10
                End If
              End If
            End If
          End If
          '
          Call MENSERR(24, "Imposible Emitir Orden de Despacho\Correspondiente a Pedido " + TRIM$(Str$(NPED&)) + ".")
          GoTo 99
          '
        End If
      End If
10  Next U&
    '
    FIN& = ULTREG&("ORDESPA")
    NODES1& = 1
    If FIN& > 0 Then
        z$ = REGLEIDO$("ORDESPA", FIN&)
        NODES1& = 1 + CVS(Left$(z$, 4))
    End If
    FF% = HOY(HOS$)
    '
    For KK1% = 1 To KKK%
        If KK1% > MAXIT% Then
            NPED& = (-1) * Abs(NPED&)
            GoTo 99
        End If
        Y$ = REGLEIDO$("ODESEMI", REODES&(KK1%))
        CI% = CVI(Mid$(Y$, 11, 2))
        CANASI = CVS(Mid$(Y$, 17, 4))
        CANASIANT = CANASI
        REGPEDA& = CVS(Mid$(Y$, 21, 4))
        
        X$ = REGLEIDO$("PEDDETA", REGPEDA&)
        If Mid$(X$, 75, 4) = Space$(4) Then
            Call REPLA(X$, MKS$(0), 75, 4)
        End If
        '
15 VDEF$ = Y$
Call REPLA(VDEF$, MKS$(CANASI), 29, 4)
OBX$ = OBJPLA$("ASIGPEDI", VDEF$)
If OBX$ = "" Then GoTo 98
If CVS(Mid$(OBX$, 29, 4)) > CANASI + 0.05 Then
   If COMALTER%("ATENCION: Está Asignando Más que el Disponible !\\Corregir\Continuar") <> 2 Then GoTo 15
End If
CANASI = CVS(Mid$(OBX$, 29, 4))
SALNOASI = CANASIANT - CANASI
If SALNOASI < 0 Then SALNOASI = 0
        '
        CASIG = CVS(Mid$(X$, 75, 4)) + CANASI
        Call REPLA(X$, MKS$(CASIG), 75, 4)
        Call GRAREG("PEDDETA", X$, REGPEDA&)
        '
        XX$ = Left$(X$, 74)
        For W& = 1 To ULTREG&("PEDDETP")
            z$ = REGLEIDO$("PEDDETP", W&)
            If Left$(z$, 30) = Left$(XX$, 30) Then
               If Mid$(z$, 33, 42) = Mid$(XX$, 33, 42) Then
                  z$ = X$
                  Call REPLA(z$, MKS$(REGPEDA&), 109, 4)
                  Call GRAREG("PEDDETP", z$, W&)
                  GoTo 20
               End If
            End If
        Next W&
        '
        MENSA$ = "Se ha Detectado un Error de Consistencia\"
        MENSA$ = MENSA$ + "en la Base de Datos de Pedidos Pendientes.\"
        MENSA$ = MENSA$ + "Correspondiente al Pedido Número " + TRIM$(Str$(Abs(NPED&))) + ".\  \"
        MENSA$ = MENSA$ + "FLEXOFT Reparará este Error oportunamente."
        Call MENSERR(24, MENSA$)
        '
20      Call REPLA(Y$, MKS$(SALNOASI), 17, 4)
        Call GRAREG("ODESEMI", Y$, REODES&(KK1%))
        '
        z$ = REGBLAN$("ORDESPA")
        Call REPLA(z$, MKS$(NODES1&), 1, 4)
        Call REPLA(z$, MKI$(FF%), 5, 2)
        Call REPLA(z$, Left$(X$, 30), 7, 30)
        Call REPLA(z$, Mid$(X$, 33, 42), 37, 42)
        Call REPLA(z$, MKS$(CANASI), 79, 4)
        Call REPLA(z$, MKS$(REGPEDA&), 111, 4)
        Call REGAPP("ORDESPA", z$)
        '
        NODES& = NODES1&
        '
98  Next KK1%
    '
99  Call CIERRARCH("ORDESPA")
    Call CIERRARCH("*.*")
    '
End Sub
'
Sub CALPESTOCE()
   '
   Screen.MousePointer = 11
   JJ& = 0
   FIN& = ULTREG&("PEDDETP")
   '
   Dim CAPE(16384), VAPE#(16384)
   '
   MODO% = COMALTER%("Clasificación por Fecha de Entrega Solicitada:\\Mes en Curso\Totales Pendientes")
   Call HEADERS("PESTOCE", "")
   If MODO% = 1 Then
         Call HEADERS("PESTOCE", "Alcance: Mes en Curso")
      ElseIf MODO% = 2 Then
         Call HEADERS("PESTOCE", "Alcance: Total de Pedidos Pendientes")
      Else
         Exit Sub
   End If
   '
   FLI1% = -1: FLI2% = FINMESACT%
   If MODO% = 2 Then
      FLI1% = -1: FLI2% = 32767
   End If
   '
   OPX% = ALTERNA%("Solo Cantidades\Cantidades e Importes")
   '
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("PEDDETP", U&)
      FENTRESOL% = CVI(Mid$(X$, 25, 2))
      If FENTRESOL% >= FLI1% Then
         If FENTRESOL% <= FLI2% Then
            CAPEN = CVS(Mid$(X$, 79, 4)) - CVS(Mid$(X$, 95, 4))
            If CAPEN < 0 Then CAPEN = 0
            If CAPEN > 0 Then
               CI% = CVI(Mid$(X$, 33, 2))
               If CI% > 0 Then
                  If CI% <= 16384 Then
                     NC% = CVI(Mid$(X$, 7, 2))
                     PREUNI = CVS(Mid$(X$, 83, 4)) * (100 - PORDESCLI(NC%)) / 100
                     If OPX% <> 2 Then PREUNI = 0
                     ITOTA# = CDbl(CAPEN * PREUNI)
                     '
                     CAPE(CI%) = CAPE(CI%) + CAPEN
                     VAPE#(CI%) = VAPE#(CI%) + ITOTA#
                  End If
               End If
            End If
         End If
      End If
   Next U&
   '
   NMS& = NUMAS%
   For U& = 1 To NMS&         '
      Call TRACE(20, U&, NMS&)
      CI% = U&
      If CAPE(CI%) >= 0.05 Then
         STACT = STODEPOS(CI%, DEPOEXPE%)
         If STACT < CAPE(CI%) Then
            FALTA = CAPE(CI%) - STACT
            Y$ = REGBLAN$("PESTOCE")
            Call REPLA(Y$, MKI$(CI%), 1, 2)
            Call REPLA(Y$, MKS$(STACT), 3, 4)
            Call REPLA(Y$, MKS$(CAPE(CI%)), 7, 4)
            Call REPLA(Y$, MKD$(VAPE#(CI%)), 11, 8)
            Call REPLA(Y$, MKS$(FALTA), 19, 4)
            Call REPLA(Y$, MKD$(VAPE#(CI%) * FALTA / CAPE(CI%)), 23, 8)
            JJ& = JJ& + 1
            Call GRAREG("PESTOCE", Y$, JJ&)
         End If
      End If
   Next U&
   '
   Call SETULTREG("PESTOCE", JJ&)
   Call CIERRARCH("PESTOCE")
   Screen.MousePointer = 1
   '
End Sub
'
Sub ANULODESPA(NODES&)
   '
   Call BLOQARCH("ORDESPA")
   Screen.MousePointer = 11
   J& = 0: FIN& = ULTREG&("ORDESPA")
   For I& = 1 To FIN&
     Call TRACE(20, I&, FIN&)
     Y$ = REGLEIDO$("ORDESPA", I&)
     If CVS(Left$(Y$, 4)) = NODES& Then
       Call REPLA(Y$, Chr$(9), 33, 1)
       Call GRAREG("ORDESPA", Y$, I&)
     End If
   Next I&
   '
   Call CIERRARCH("*.*")
   Call LIBARCH("ORDESPA")
   Call BAJALDISCO
   Screen.MousePointer = 1
   '
End Sub
'
Sub CONODESCLI()
   '
   Dim NUPEDSEL&(1024)
   '
   Call SELECHO("DEUDOR1")
   NCXI% = Val(TRIM$(VALACT1$("DEUDOR1")))
   If NCXI% < 1 Then Exit Sub
   '
   RPEDX$ = RECFILT$("ORDESPA", 5, MKI$(NCXI%))
   NRODESA& = 0
   FIN& = Len(RPEDX$)
   For UU& = 1 To Len(RPEDX$) Step 4
     U& = CVS(Mid$(RPEDX$, UU&, 4))
     If FIN& > 0 Then
        Call TRACE(20, UU&, FIN&)
     End If
     X$ = REGLEIDO$("ORDESPA", U&)
     If CVI(Mid$(X$, 13, 2)) = NCXI% Then
       CAN = CVS(Mid$(X$, 79, 4))
       CADESPACH = CVS(Mid$(X$, 83, 4))
       CAPENDE = CAN - CADESPACH
       If CAN >= 0.05 Then
         '
         NRODES& = CVS(Left$(X$, 4))
         If NRODES& <> NRODESA& Then
            TTXX$ = AJUSTI$("OD-" + TRIM$(CSTRING$(MKS$(NRODES&), 3, 8, 0, 2)), 11)
            FEX = CVI(Mid$(X$, 31, 2))
            TTXX$ = TTXX$ + "   Sol." + FECHATEX$(FEX)
            FORSELPED.LIPEDPE.AddItem TTXX$
         End If
         NRODESA& = NRODES&
         CAPED% = CAPED% + 1
         FENTRE% = CVI(Mid$(X$, 31, 2))
         If FENTRE% < FINF Then FINF = FENTRE%
         If FENTRE% > FSUP Then FSUP = FENTRE%
       End If
     End If
   Next UU&
   '
   FORSELPED.Show 1
   '
   CAPESEL% = 0
   PRESEFA.Label4 = ""
   For U& = 1 To FORSELPED.LIPEDPE.ListCount
     If FORSELPED.LIPEDPE.Selected(U& - 1) = True Then
       CAPESEL% = CAPESEL% + 1
       TTXX$ = FORSELPED.LIPEDPE.List(U& - 1)
       NUPEDSEL&(CAPESEL%) = Val(Left$(TTXX$, 10))
       If Left$(TTXX$, 3) = "OD-" Then
           NUPEDSEL&(CAPESEL%) = Val(Mid$(TTXX$, 4, 10))
           PRESEFA.Label4.Caption = TRIM$(PRESEFA.Label4.Caption + " " + TRIM$(Left$(TTXX$, 10)))
         ElseIf Left$(TTXX$, 3) = "RT-" Then
           NUPEDSEL&(CAPESEL%) = (-1) * Val(Mid$(TTXX$, 4, 10))
       End If
     End If
   Next U&
   '
   If CAPESEL% < 1 Then Exit Sub
   '
   LPRE% = LIPRCLI%(NCXI%)
   If LPRE% < 1 Then LPRE% = 1
   ARCHIPRE$ = TRIM$(Str$(LPRE%))
   While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
   ARCHIPRE$ = "LIPRE" + ARCHIPRE$
   '
   CAITCAR% = 0
   PRESEFA.List1.Clear
   PRESEFA.Label1.Caption = ""
   FIN& = Len(RPEDX$)
   SUMATO# = 0
   For UU& = 1 To Len(RPEDX$) Step 4
     U& = CVS(Mid$(RPEDX$, UU&, 4))
     If FIN& > 0 Then
        Call TRACE(20, UU&, FIN&)
     End If
     '
     REGPEDP& = U&
     REGODES& = U&
     X1$ = REGLEIDO$("ORDESPA", U&)
     REGPEDA& = CVS(Mid$(X1$, 111, 4))
     X2$ = REGLEIDO$("PEDDETA", REGPEDA&)
     X$ = X2$
     '
     If CVI(Mid$(X$, 7, 2)) = NCXI% Then
       STATUS% = Asc(Mid$(X$, 27, 1))
       If STATUS% <= 1 Then
          NODES& = CVS(Left$(X1$, 4))
          For UI% = 1 To CAPESEL%
            If NUPEDSEL&(UI%) = NODES& Then GoTo 20
          Next UI%
       End If
     End If
     GoTo 39
     '
20   CAN = CVS(Mid$(X$, 79, 4)) - CVS(Mid$(X$, 95, 4))
     If CAN >= 0.05 Then
       CI% = CVI(Mid$(X$, 33, 2))
       z$ = Space$(128)
       Call REPLA(z$, CSTRING$(MKS$(CAN), 3, 5, 0, 2), 1, 5)  ' cantidad
       Call REPLA(z$, CODEXT$(CI%), 8, 12)
       Call REPLA(z$, DESCRIT$(CI%), 21, 40)
         PRELI = CVS(Mid$(X$, 71, 4))
         If PRELI < 0.0001 Then PRELI = PRELISTA#(ARCHIPRE$, CI%)
         PRECIO = CVS(Mid$(X$, 83, 4))
         If PRELI > 0.005 And PRELI < PRECIO Then PRELI = PRECIO
       Call REPLA(z$, CSTRING$(MKS$(PRELI), 3, 10, 4, 2), 63, 10)
         IMPO# = CDbl(PRELI) * CAN
       Call REPLA(z$, CSTRING$(MKD$(IMPO#), 3, 12, 2, 2), 73, 12)
       SUMATO# = SUMATO# + IMPO#
       PRESEFA.List1.AddItem z$
     End If
     '
39 Next UU&
   '
   PRESEFA.Caption = RASOCLI$(NCXI%)
   PRESEFA.Label1.Caption = CSTRING$(MKD$(SUMATO#), 3, 12, 2, 2)
   '
   PRESEFA.Show 1
   
End Sub

