VERSION 5.00
Begin VB.Form ODESPACHO 
   BackColor       =   &H00E0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asignación de Pedidos - Ordenes de Despacho"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5595
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
   ScaleHeight     =   6540
   ScaleWidth      =   5595
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   4560
      ScaleHeight     =   75
      ScaleWidth      =   810
      TabIndex        =   27
      Top             =   2460
      Width           =   870
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   28
         Top             =   0
         Width           =   12000
      End
   End
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
      Picture         =   "ODESPACHO.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1320
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0FF&
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
         TabIndex        =   23
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
            Picture         =   "ODESPACHO.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   24
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
            TabIndex        =   25
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
         TabIndex        =   20
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
            Picture         =   "ODESPACHO.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   21
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
            TabIndex        =   22
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
            Left            =   0
            Picture         =   "ODESPACHO.frx":0C7E
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
            Picture         =   "ODESPACHO.frx":0F88
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
            Left            =   0
            Picture         =   "ODESPACHO.frx":10D2
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
      BackColor       =   &H00E0E0FF&
      Caption         =   "Consultas - Listados"
      Height          =   2790
      Left            =   240
      TabIndex        =   2
      Top             =   3600
      Width           =   5160
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0FF&
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
         TabIndex        =   26
         Top             =   840
         Width           =   4530
      End
      Begin VB.OptionButton Option7 
         BackColor       =   &H00E0E0FF&
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
         TabIndex        =   19
         Top             =   1470
         Width           =   4275
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
         Picture         =   "ODESPACHO.frx":13DC
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2100
         Width           =   855
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00E0E0FF&
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
         BackColor       =   &H00E0E0FF&
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
            BackColor       =   &H00E0E0FF&
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
            BackColor       =   &H00E0E0FF&
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
         BackColor       =   &H00E0E0FF&
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
      Picture         =   "ODESPACHO.frx":16E6
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Doble Click s/Flexoft para Recuperar Indices"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2205
      TabIndex        =   29
      Top             =   0
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Line Line1 
      X1              =   4350
      X2              =   5625
      Y1              =   3350
      Y2              =   3350
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   3990
      Picture         =   "ODESPACHO.frx":1830
      ToolTipText     =   "Haga Doble-Click para Refrescar Indices"
      Top             =   2830
      Width           =   2010
   End
End
Attribute VB_Name = "ODESPACHO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Private Sub Command2_Click()
    '
    Command2.Enabled = False
    '
    If Option1.Value = True Then
       If Option5.Value = True Then
            Call CONECRUN("*CODESEM", "")
          Else
            Call CONECRUN("*ODESEMI", "")
       End If
       GoTo 99
    End If
    '
    If Option2.Value = True Then
       Call CONODESCLI
       GoTo 99
    End If
    '
    If Option6.Value = True Then
       Call ODENCURNU      ' ordenes de despacho en curso por numero
       If ULTREG&("ORDESPE") > 0 Then
         If Option5.Value = True Then
              Call CONECRUN("*COODESP", "")
            Else
              Call CONECRUN("*LIODESP", "")
         End If
       End If
       GoTo 99
    End If
    '
    If Option7.Value = True Then
       Call CALPESTOCE
       If Option5.Value = True Then
            Call CONECRUN("*SOSTOCC", "")
          Else
            Call CONECRUN("*SOSTOCE", "")
       End If
       GoTo 99
    End If
    '
99  Command2.Enabled = True
    '
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    Call CALASIAUT
    Command3.Enabled = True
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
    Call HELPONLINE("ODESPACH")
End Sub

Private Sub Command7_Click()
    Command7.Enabled = False
    Screen.MousePointer = 11
    JJ& = 0
    RFF$ = RECFILT$("MAVENDID", 17, Chr$(1)) ' solo pendientes
    For UK& = 1 To Len(RFF) Step 4
        U& = CVS(Mid$(RFF$, UK&, 4))
        X$ = REGLEIDO$("MAVENDID", U&)
        NODES& = CVS(Mid$(X$, 87, 4))
        If NODES& > 0 Then
           If NODES& <> NODESA& Then
              NODESA& = NODES&
              Z$ = REGBLAN$("INDIODES")
              NC% = CVI(Mid$(X$, 1, 2))
              CLI$ = "": If NC% > 0 Then CLI$ = RASOCLI$(NC%)
              FEX = CVI(Mid$(X$, 81, 2))
              Z1$ = FECHATEX$(FEX) + "  " + CLI$
              '
              Call REPLA(Z$, MKS$(NODES&), 1, 4)
              Call REPLA(Z$, Z1$, 5, 40)
              JJ& = JJ& + 1
              Call GRAREG("INDIODES", Z$, JJ&)
           End If
        End If
    Next UK&
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
   RFF$ = RECFILT$("MAVENDID", 17, Chr$(1))
   FIN& = Len(RFF$)
   JJ& = 0
   
   For UK& = 1 To FIN& Step 4
     '
     Call TRACE(20, UK&, FIN&)
     U& = CVS(Mid$(RFF$, UK&, 4))
     X$ = REGLEIDO$("MAVENDID", U&)
     '
     NODES& = CVS(Mid$(X$, 87, 4))
     If NODES& > 0 Then
       If NODES& <> NODESA& Then
         NODESA& = NODES&
         Z$ = REGBLAN$("INDIODES")
         NC% = CVI(Mid$(X$, 1, 2))
         CLI$ = "": If NC% > 0 Then CLI$ = RASOCLI$(NC%)
         FEX = CVI(Mid$(X$, 81, 2))
         Z1$ = FECHATEX$(FEX) + "  " + CLI$
         '
         Call REPLA(Z$, MKS$(NODES&), 1, 4)
         Call REPLA(Z$, Z1$, 5, 40)
         JJ& = JJ& + 1
         Call GRAREG("INDIODES", Z$, JJ&)
       End If
     End If
   Next UK&
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
     Call FRESHECHO("INDIODES")
     GoTo 5
   End If
   '
   Command8.Enabled = True
   '
End Sub

Private Sub Command9_Click()
For UI% = 1 To NUMAS%
STA = STOCKASI(UI%)
If STA > 0 Then
If STODEPOS(UI%, DEPOEXPE%) < STOCKASI(UI%) Then
MsgBox CODEXT$(UI%)
End If
End If
Next UI%

End Sub

Private Sub Form_Load()
   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then Close: End
   '
   Option1.Value = True
   Show
   Refresh
   '
   If ULTREG&("MAVENDID") < 1 Then
     Call CONVERTIR
   End If
   '
End Sub

Sub CALASIAUT()          ' revision de situacion de stocks
    '                      para asignacion de pedidos
    Screen.MousePointer = 11
    Dim PUNPED$(16384)
    Dim STOLE%(16384)
    Dim STTOT(16384)
    Dim STASI(16384)
    '
    HOI% = HOY(HOS$)
    DIANTIDES% = Val(CONTROL$("", "ANTIDESP"))
    If DIANTIDES% < 1 Then DIANTIDES% = 3
    '
    ' revisar y depurar pedidos pendientes
    '
    IFIX$ = LOADFILE$(LUDAT$ + "PEDDETA.X07")
    RFF$ = ""
    For IQQ& = 1 To Len(IFIX$)
      STATI% = Asc(Mid$(IFIX$, IQQ&, 1))
      If STATI% = 1 Or STATI% = 2 Then
        RFF$ = RFF$ + MKS$(IQQ&)
      End If
    Next IQQ&
    'RFF$ = RECFILT$("PEDDETA", 7, Chr$(1)) ' solo pendientes
    FIREFI& = Len(RFF$): V1& = 0
    '
    For UU& = 1 To FIREFI& Step 4
      '
      U& = CVS(Mid$(RFF$, UU&, 4))
      Call TRACE(20, UU&, FIREFI&)
      X$ = REGLEIDO$("PEDDETA", U&)
      '
      If Asc(Mid$(X$, 27, 1)) <> 9 Then ' no esta anulado
        '
        CAPEDIDA = CVS(Mid$(X$, 79, 4))
        If CAPEDIDA < 0.05 Then       ' marca como anulado
          Call REPLA(X$, Chr$(9), 27, 1)
          Call GRAREG("PEDDETA", X$, U&)
        End If
        '
        If CAPEDIDA >= 0.05 Then
          '
          CADESPACH = CVS(Mid$(X$, 95, 4))
          CAPENDENT = CAPEDIDA - CADESPACH
          '
          If Asc(Mid$(X$, 27, 1)) < 3 Then
            If CAPENDENT < 0.05 Then   ' marca como entregado
              Call REPLA(X$, Chr$(3), 27, 1)
              Call GRAREG("PEDDETA", X$, U&)
            End If
          End If
          '
          If CAPENDENT >= 0.05 Then
            '
            NCLIE% = CVI(Mid$(X$, 7, 2))
            NPED& = CVS(Left$(X$, 4))
            '
            If NPED& > 0 Then
              '
              CI% = CVI(Mid$(X$, 33, 2))
              REGENCA& = CVS(Mid$(X$, 109, 4))
              '
              Z$ = REGLEIDO$("PEDENCA", REGENCA&)
              NPEDA& = CVS(Left$(Z$, 4))
              NCLIA% = CVI(Mid$(Z$, 55, 2))
              '
              If NPEDA& <> NPED& Or NCLIA% <> NCLIE% Then
                Call MENSERR(24, "Punteros Incorrectos en Pedido " + TRIM$(Str$(NPED&)) + "/" + TRIM$(Str$(NPEDA&)) + "\Consulte a su Soporte Técnico")
                GoTo 90
              End If
              '
              If Asc(Mid$(Z$, 267, 1)) > 1 Then
                Call REPLA(X$, Chr$(9), 27, 1)  ' anula detalle
                Call GRAREG("PEDDETA", X$, U&)
                GoTo 90
              End If
              '
              Call REPLA(X$, MKS$(0), 75, 4)   ' pone en cero cantidad asignada
              Call GRAREG("PEDDETA", X$, U&)
              '
              Y$ = X$
              Call REPLA(Y$, MKS$(0), 75, 4)   ' pone en cero cantidad asignada
              Call REPLA(Y$, MKS$(U&), 109, 4) ' puntero a PEDDETA
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
90  Next UU&
    '
    Call SETULTREG("PEDDETP", V1&)
    Call CIERRARCH("PEDDETP")
    Call CIERRARCH("PEDDETA")
    '
    ' depurar segun ordenes de despacho en curso
    '
    UREPEP& = ULTREG&("PEDDETP")
    RFF$ = RECFILT$("MAVENDID", 17, Chr$(1)) ' solo pendientes
    FIREFI& = Len(RFF$): V1& = 0
    '
    For UU& = 1 To FIREFI& Step 4
      '
      U& = CVS(Mid$(RFF$, UU&, 4))
      Call TRACE(20, UU&, FIREFI&)
      X$ = REGLEIDO$("MAVENDID", U&)
      '

      NODE& = CVS(Mid$(X$, 87, 4))
      NPED& = CVS(Mid$(X$, 93, 4))
      '
      If NPED& > 0 Then
        CADES = CVS(Mid$(X$, 177, 4))
        CAENT = CVS(Mid$(X$, 225, 4))
        '
        If CAENT >= 0.05 Then
          Call REPLA(X$, Chr$(2), 63, 1) ' marca como entregada
          Call REPLA(X$, MKI$(0), 65, 2) ' pone en cero cliente
          Call REPLA(X$, MKI$(0), 67, 2) ' pone en cero articulo
          Call GRAREG("MAVENDID", X$, U&)
          CADES = 0
        End If
        '
        If CADES >= 0.05 Then
          CI% = CVI(Mid$(X$, 5, 2))
          If CI% > 0 Then
            If CI% <= NUMAS% Then
              REGPEDA& = CVS(Mid$(X$, 99, 4))
              For II% = 1 To Len(PUNPED$(CI%)) Step 4
                REGPEDP& = CVS(Mid$(PUNPED$(CI%), II%, 4))
                If REGPEDP& > 0 Then
                  If REGPEDP& <= UREPEP& Then
                    Y$ = REGLEIDO$("PEDDETP", REGPEDP&)
                    NPEDA& = CVS(Left$(Y$, 4))
                    If NPEDA& = NPED& Then
                      If CVS(Mid$(Y$, 109, 4)) = REGPEDA& Then
                        '
                        TOASI = CVS(Mid$(Y$, 75, 4)) + CADES
                        Call REPLA(Y$, MKS$(TOASI), 75, 4)
                        Call GRAREG("PEDDETP", Y$, REGPEDP&)
                        '
                        Y$ = REGLEIDO$("PEDDETA", REGPEDA&)
                        Call REPLA(Y$, MKS$(TOASI), 75, 4)
                        Call GRAREG("PEDDETA", Y$, REGPEDA&)
                        '
                        GoTo 95
                        '
                      End If
                    End If
                  End If
                End If
              Next II%
              '
' ATENCION - Las que llegan aqui es porque la O/D figura pendiente de
'            entrega pero el pedido esta saldado - son pocos casos pero
'            hay que revisar oportunamente
              '
              Y$ = REGLEIDO$("PEDDETA", REGPEDA&)
              If CVS(Left$(Y$, 4)) = NPED& Then
                If CVI(Mid$(Y$, 7, 2)) = CVI(Left$(X$, 2)) Then  'coincide cliente
                  If CVI(Mid$(Y$, 33, 2)) = CVI(Mid$(X$, 5, 2)) Then ' COINCIDE ARTICULO
                    GoTo 95
                  End If
                End If
              End If
              '
              Call MENSERR(24, "No se Encontró Pedido Original " + TRIM$(Str$(NPED&)) + "\Correspondiente a la Orden de Despacho Nro." + Str$(NODE&) + ".")
              '
95            If STOLE%(CI%) = 0 Then
                STOLE%(CI%) = 1
                STTOT(CI%) = STODEPOS(CI%, DEPOEXPE%)
              End If
              STTOT(CI%) = STTOT(CI%) - CADES
              '
            End If
          End If
          '
        End If
      End If
      '
99  Next UU&
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
100             Call REPLA(Y$, MKS$(CANASI), 17, 4)
                Call REPLA(Y$, MKS$(REGPEDA&), 21, 4)   ' registro PEDDETA
                Call REPLA(Y$, MKS$(U&), 29, 4) ' registro PEDDETP
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
    '
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
            Z$ = CSTRING$(MKS$(NPED&), 3, 7, 0, 2) + "  " + RASOCLI$(NC%)
            FODESEMI.LIPEDES.AddItem Z$
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
        CAPESEL% = CAPESEL% + 1
        TTXX$ = FODESEMI.LIPEDES.List(U& - 1)
        '
        NPED& = Val(Left$(TTXX$, 7))
20      Call GENODESPA(NPED&, NODES&, 2, OKEY%)
        If OKEY% < 0 Then
           Call COMUNI("Proceso de Asignación Cancelado")
           GoTo 90      ' CANCELA PROCESO
        End If
        If OKEY% = 1 Then
          Call PRIODESPA(NODES&)
        End If
      End If
    Next U&
    '
90  Call CIERRARCH("*.*")
    Call BAJALDISCO
    '
    Screen.MousePointer = 1
    '
End Sub
'
Sub CALPESTOCE()
   '
   Screen.MousePointer = 11
   JJ& = 0
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
   FIN& = ULTREG&("PEDDETP")
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
   ANULAX% = 0
   Call BLOQARCH("MAVENDID")
   Screen.MousePointer = 11
   RFF$ = RECFILT$("MAVENDID", 24, MKS$(NODES&))
   FIN& = Len(RFF$)
   '
   For U& = 1 To FIN& Step 4
     Call TRACE(20, U&, FIN&)
     I& = CVS(Mid$(RFF$, U&, 4))
     Y$ = REGLEIDO$("MAVENDID", I&)
     If CVS(Mid$(Y$, 87, 4)) = NODES& Then
       If Asc(Mid$(Y$, 63, 1)) = 1 Then
         Call REPLA(Y$, Chr$(9), 63, 1)
         Call REPLA(Y$, MKI$(0), 65, 2)
         Call REPLA(Y$, MKI$(0), 67, 2)
         '
         FANU% = HOY(HOS$)
         HANU% = HORANUM%(Time$)
         Call REPLA(Y$, MKI$(FANU%), 187, 2)
         Call REPLA(Y$, MKI$(HANU%), 189, 2)
         Call REPLA(Y$, Chr$(USNBR% Mod 100), 191, 1)
         '
         Call GRAREG("MAVENDID", Y$, I&)
         ANULAX% = 1
         '
         ' ANULAR DESPACHO EN PEDDETA Y PEDDETP
         '
         NPEDID& = CVS(Mid$(Y$, 93, 4))
         REGPEDA& = CVS(Mid$(Y$, 99, 4))
         NCLIE% = CVI(Mid$(Y$, 1, 2))
         CIXI% = CVI(Mid$(Y$, 5, 2))
         If REGPEDA& > 0 Then
           If REGPEDA& <= ULTREG&("PEDDETA") Then
             X$ = REGLEIDO$("PEDDETA", REGPEDA&)
             If CVS(Left$(X$, 4)) = NPEDID& Then
               If CVI(Mid$(X$, 7, 2)) = NCLIE% Then
                 If CVI(Mid$(X$, 33, 2)) = CIXI% Then
                   If Asc(Mid$(X$, 27, 1)) <= 2 Then
                     Call REPLA(X$, MKS$(0), 75, 4) ' ANULA CANTIDAD ASIGNADA
                     Call REPLA(X$, Chr$(1), 27, 1) ' STATUS PASA A 1
                     Call GRAREG("PEDDETA", X$, REGPEDA&)
                   End If
                 End If
               End If
             End If
             '
             For UZ& = 1 To ULTREG&("PEDDETP")
               X$ = REGLEIDO$("PEDDETP", UZ&)
               If CVS(Left$(X$, 4)) = NPEDID& Then
                 If CVI(Mid$(X$, 7, 2)) = NCLIE% Then
                   If CVI(Mid$(X$, 33, 2)) = CIXI% Then
                     If CVS(Mid$(X$, 109, 4)) = REGPEDA& Then
                       Call REPLA(X$, MKS$(0), 75, 4) ' ANULA CANTIDAD ASIGNADA
                       Call REPLA(X$, Chr$(1), 27, 1) ' STATUS PASA A 1
                       Call GRAREG("PEDDETP", X$, UZ&)
                     End If
                   End If
                 End If
               End If
             Next UZ&
             '
           End If
         End If
       End If
     End If
   Next U&
   '
   Call CIERRARCH("MAVENDID")
   Call CIERRARCH("PEDDETA")
   Call CIERRARCH("PEDDETP")
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Screen.MousePointer = 1
   '
End Sub
'
Sub ODENCURNU()     ' ordenes de despacho en curso por numero
   '
   Screen.MousePointer = 11
   '
   REBLA$ = REGBLAN$("ORDESPE")
   RFF$ = RECFILT$("MAVENDID", 17, Chr$(1))
   FIN& = Len(RFF$)
   JJ& = 0
   For UL& = 1 To FIN& Step 4
     Call TRACE(20, UL&, FIN&)
     U& = CVS(Mid$(RFF$, UL&, 4))
     X$ = REGLEIDO$("MAVENDID", U&)
     NODES& = CVS(Mid$(X$, 87, 4))
     If NODES& > 0 Then
       NPED& = CVS(Mid$(X$, 93, 4))
       If NPED& > 0 Then
         CAPEN = CVS(Mid$(X$, 177, 4)) - CVS(Mid$(X$, 225, 4))
         If CAPEN >= 0.05 Then
           CI% = CVI(Mid$(X$, 5, 2))
           If CI% > 0 Then
             If CI% <= NUMAS% Then
               FODES% = CVI(Mid$(X$, 81, 2))
               FEPED% = CVI(Mid$(X$, 103, 2))
               NCLIE% = CVI(Mid$(X$, 1, 2))
               NOCOM$ = TRIM$(Mid$(X$, 105, 16))
               FESOL% = CVI(Mid$(X$, 125, 2))
               LICOVE$ = Mid$(X$, 121, 3)
               COTAMED$ = Mid$(X$, 9, 36)
               PRELIS = CVD(Mid$(X$, 127, 8))
               CASIG = CVS(Mid$(X$, 177, 4))
               CAENT = CVS(Mid$(X$, 225, 4))
               SALPE = CASIG - CAENT: If SALPE < 0 Then SALPE = 0
               FEREM% = CVI(Mid$(X$, 199, 2))
               NUREM = CVS(Mid$(X$, 203, 4))
               '
               Y$ = REBLA$
               Call REPLA(Y$, MKS$(NODES&), 1, 4)
               Call REPLA(Y$, MKI$(FODES%), 5, 2)
               Call REPLA(Y$, MKS$(NPED&), 7, 4)
               Call REPLA(Y$, MKI$(FEPED%), 11, 2)
               Call REPLA(Y$, MKI$(NCLIE%), 13, 2)
               Call REPLA(Y$, NOCOM$, 15, 16)
               Call REPLA(Y$, MKI$(FESOL%), 31, 2)
               Call REPLA(Y$, Chr$(1), 33, 1)
               Call REPLA(Y$, LICOVE$, 34, 3)
               Call REPLA(Y$, MKI$(CI%), 37, 2)
               Call REPLA(Y$, COTAMED$, 39, 36)
               Call REPLA(Y$, MKS$(PRELIS), 75, 4)
               Call REPLA(Y$, MKS$(CASIG), 79, 4)
               Call REPLA(Y$, MKS$(CAENT), 83, 4)
               Call REPLA(Y$, MKS$(SALPE), 87, 4)
               Call REPLA(Y$, MKI$(FEREM), 91, 2)
               Call REPLA(Y$, MKS$(NUREM), 99, 4)
               '
               JJ& = JJ& + 1
               Call GRAREG("ORDESPE", Y$, JJ&)
             End If
           End If
         End If
       End If
     End If
   Next UL&
   '
   Call SETULTREG("ORDESPE", JJ&)
   Call CIERRARCH("ORDESPE")
   Call CIERRARCH("MAVENDID")
   Screen.MousePointer = 1
   '
   If JJ& < 1 Then
     Call COMUNI("No Hay Ordenes de Despacho en Curso")
   End If
   '
End Sub

Sub CONVERTIR()
   '
   'Este Modulo convierte las viejas bases de datos
   'ORDESPB y ORDESPA al nuevo formato MAVENDID
   '
   If COMALTER%("FLEXOFT Re-organizará Ahora la Base de Datos\de Ordenes de Despacho.\  \Cierre las Aplicaciones de Pedidos y Facturación.\\Cancelar\Continuar") <> 2 Then
     Call CIERRARCH("*.*")
     Call FINAL("")
   End If
   '
   Call BLOQARCH("MAVENDID")
   Call BLOQARCH("ORDESPB")
   Call BLOQARCH("ORDESPA")
   '
   Screen.MousePointer = 11
   UREPE& = ULTREG&("PEDDETA")
   REBLA$ = REGBLAN$("MAVENDID")
   '
   'ojo verificar los seteos especificos para sabo
   '
   FIN& = ULTREG&("ORDESPB")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     X$ = REGLEIDO$("ORDESPB", U&)
     NODES& = CVS(Left$(X$, 4))
     FEDES% = CVI(Mid$(X$, 5, 2))
     NOPED& = CVS(Mid$(X$, 7, 4))
     FEPED% = CVI(Mid$(X$, 11, 2))
     NCLIE% = CVI(Mid$(X$, 13, 2))
     NUOCO$ = AJUSTI$(Mid$(X$, 15, 16), 16)
     FENTRESO% = CVI(Mid$(X$, 31, 2))
     STATUS% = 2                     ' ENTREGADO
     PRECOVEN$ = Mid$(X$, 34, 3) 'L.PRECIO-C.PAGO-VENDEDOR
     CIXI% = CVI(Mid$(X$, 37, 2))
     KCOLOR$ = Mid$(X$, 39, 16)
     KTALLE$ = Mid$(X$, 55, 8)
     KLARGO = CVS(Mid$(X$, 63, 4))
     KANCHO = CVS(Mid$(X$, 67, 4))
     KESPES = CVS(Mid$(X$, 71, 4))
     PREUNID# = CVS(Mid$(X$, 75, 4))
     CANASI = CVS(Mid$(X$, 79, 4))
     CAENTR = CVS(Mid$(X$, 83, 4))
     FENTRE% = CVI(Mid$(X$, 91, 2))
     NREMI& = CVS(Mid$(X$, 99, 4))
     REPEA& = CVS(Mid$(X$, 111, 4))
     '
     If REPEA& > 0 Then
       If REPEA& <= UREPE& Then
         Z$ = REGLEIDO$("PEDDETA", REPEA&)
         NPP& = CVS(Left$(Z$, 4))
         If NPP& = NOPED& Then
           FPP% = CVI(Mid$(Z$, 5, 2))
           If FPP% = FEPED% Then
             NCC% = CVI(Mid$(Z$, 7, 2))
             If NCC% = NCLIE% Then
               CKKI% = CVI(Mid$(Z$, 33, 2))
               If CKKI% = CIXI% Then
                 PORCOVE = CVI(Mid$(Z$, 31, 2)) / 10
                 PORCOVE = 0 ' PARA SABO
               End If
             End If
           End If
         End If
       End If
     End If
     '
     Y$ = REBLA$
     Call REPLA(Y$, MKI$(NCLIE%), 1, 2)
     Call REPLA(Y$, Chr$(0), 3, 1) ' codigo de empresa
     Call REPLA(Y$, Chr$(0), 4, 1) ' CODIGO DE SUCURSAL
     Call REPLA(Y$, MKI$(CIXI%), 5, 2)
     Call REPLA(Y$, AJUSTI$(KCOLOR$, 16), 7, 16)
     Call REPLA(Y$, AJUSTI$(KTALLE$, 8), 25, 8)
     Call REPLA(Y$, MKS$(KLARGO), 33, 4)
     Call REPLA(Y$, MKS$(KANCHO), 37, 4)
     Call REPLA(Y$, MKS$(KESPES), 41, 4)
     Call REPLA(Y$, String$(4, 0), 45, 4)  ' FILLER
     UNI$ = UNIMED$(CIXI%)
     Call REPLA(Y$, UNI$, 49, 4)
     Call REPLA(Y$, UNI$, 53, 4)
     Call REPLA(Y$, MKS$(1), 57, 4)  ' COEFICIENTE CONVERSION UNIDADES
     Call REPLA(Y$, "PC", 61, 2)     ' MOVIMIENTO DE ORIGEN
     Call REPLA(Y$, Chr$(2), 63, 1)  ' STATUS ENTREGADO
     Call REPLA(Y$, Chr$(0), 64, 1)  ' FILLER
     Call REPLA(Y$, MKI$(0), 65, 2)  ' CLIENTE
     Call REPLA(Y$, MKI$(0), 67, 2)  ' CODIGO
     Call REPLA(Y$, String$(12, 0), 69, 12) ' FILLER
     '
     Call REPLA(Y$, MKI$(FEDES%), 81, 2) ' FECHA O/DESPACHO
     Call REPLA(Y$, String$(4, 0), 83, 4) ' HORA-USUA-TERM
     Call REPLA(Y$, MKS$(NODES&), 87, 4)  ' NRO O/DESP
     Call REPLA(Y$, MKI$(0), 91, 2)       ' NRO DE ORDEN
     Call REPLA(Y$, MKS$(NOPED&), 93, 4)
     Call REPLA(Y$, Chr$(0), 97, 1) ' codigo de empresa
     Call REPLA(Y$, Chr$(0), 98, 1) ' CODIGO DE SUCURSAL
     Call REPLA(Y$, MKS$(REPEA&), 99, 4) ' REGISTRO PEDDETA
     Call REPLA(Y$, MKI$(FEPED%), 103, 2) ' FECHA DE PEDIDO
     Call REPLA(Y$, AJUSTI$(NUOCO$, 16), 105, 16)
     Call REPLA(Y$, PRECOVEN$ + Chr$(0), 121, 4)
     Call REPLA(Y$, MKI$(FENTRESO%), 125, 2) ' FECHA O/DESPACHO
     Call REPLA(Y$, MKD$(PREUNID#), 127, 8) ' PRECIO UNITARIO BRUTO
     Call REPLA(Y$, MKS$(0), 135, 4) ' PORCIENTO DESCUENTO item
     Call REPLA(Y$, MKD$(PREUNID#), 139, 8) ' PRECIO UNITARIO NETO
     Call REPLA(Y$, MKS$(0), 147, 4) ' PORCIENTO DESCUENTO general pedido
     Call REPLA(Y$, MKD$(PREUNID#), 151, 8) ' PRECIO UNITARIO NETO
     Call REPLA(Y$, Chr$(0), 159, 1)       ' MONEDA = $
     Call REPLA(Y$, Chr$(0), 160, 1)       ' FILLER
     Call REPLA(Y$, MKD$(1), 161, 8)       ' TIPO DE CAMBIO
     Call REPLA(Y$, MKD$(PREUNID#), 169, 8) ' valor unitario neto-neto EN PESOS
     Call REPLA(Y$, MKS$(CANASI), 177, 4)  ' CANTIDAD ASIGNADA
     Call REPLA(Y$, MKS$(PORCOVE), 181, 4)  ' POCEN.COMIS.VENDEDOR
     Call REPLA(Y$, String$(2, 0), 185, 2) ' FILLER
     Call REPLA(Y$, String$(6, 0), 187, 6) ' ANULACION (F/H/U/T)
     '
     Call REPLA(Y$, MKI$(FENTRE%), 193, 2) ' FECHA ENTREGA (REMITO)
     Call REPLA(Y$, String$(4, 0), 195, 4) ' HORA-USUA-TERM
     Call REPLA(Y$, MKI$(FENTRE%), 199, 2) ' FECHA ENTREGA (REMITO)
     Call REPLA(Y$, String$(4, 0), 83, 4) ' HORA-USUA-TERM
     Call REPLA(Y$, MKS$(NREMI&), 203, 4)  ' NRO REMITO
     Call REPLA(Y$, String$(2, 0), 207, 2) ' EMPRESA / SUCURSAL REMITO
' FALTA NUMERO COMPLETO REMITO
     Call REPLA(Y$, MKS$(CAENTR), 225, 4)  ' CANTIDAD ASIGNADA
     '
     Call REGAPP("MAVENDID", Y$)
     '
99 Next U&
     
   FIN& = ULTREG&("ORDESPA")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     X$ = REGLEIDO$("ORDESPA", U&)
     NODES& = CVS(Left$(X$, 4))
     FEDES% = CVI(Mid$(X$, 5, 2))
     NOPED& = CVS(Mid$(X$, 7, 4))
     FEPED% = CVI(Mid$(X$, 11, 2))
     NCLIE% = CVI(Mid$(X$, 13, 2))
     NUOCO$ = AJUSTI$(Mid$(X$, 15, 16), 16)
     FENTRESO% = CVI(Mid$(X$, 31, 2))
     STATUS% = 1                     ' PENDIENTE
     PRECOVEN$ = Mid$(X$, 34, 3) 'L.PRECIO-C.PAGO-VENDEDOR
     CIXI% = CVI(Mid$(X$, 37, 2))
     KCOLOR$ = Mid$(X$, 39, 16)
     KTALLE$ = Mid$(X$, 55, 8)
     KLARGO = CVS(Mid$(X$, 63, 4))
     KANCHO = CVS(Mid$(X$, 67, 4))
     KESPES = CVS(Mid$(X$, 71, 4))
     PREUNID# = CVS(Mid$(X$, 75, 4))
     CANASI = CVS(Mid$(X$, 79, 4))
     CAENTR = CVS(Mid$(X$, 83, 4))
     FENTRE% = CVI(Mid$(X$, 91, 2))
     NREMI& = CVS(Mid$(X$, 99, 4))
     REPEA& = CVS(Mid$(X$, 111, 4))
     '
     If REPEA& > 0 Then
       If REPEA& <= UREPE& Then
         Z$ = REGLEIDO$("PEDDETA", REPEA&)
         NPP& = CVS(Left$(Z$, 4))
         If NPP& = NOPED& Then
           FPP% = CVI(Mid$(Z$, 5, 2))
           If FPP% = FEPED% Then
             NCC% = CVI(Mid$(Z$, 7, 2))
             If NCC% = NCLIE% Then
               CKKI% = CVI(Mid$(Z$, 33, 2))
               If CKKI% = CIXI% Then
                 PORCOVE = CVI(Mid$(Z$, 31, 2)) / 10
                 PORCOVE = 0 ' PARA SABO
               End If
             End If
           End If
         End If
       End If
     End If
     '
     Y$ = REBLA$
     Call REPLA(Y$, MKI$(NCLIE%), 1, 2)
     Call REPLA(Y$, Chr$(0), 3, 1) ' codigo de empresa
     Call REPLA(Y$, Chr$(0), 4, 1) ' CODIGO DE SUCURSAL
     Call REPLA(Y$, MKI$(CIXI%), 5, 2)
     Call REPLA(Y$, AJUSTI$(KCOLOR$, 16), 7, 16)
     Call REPLA(Y$, AJUSTI$(KTALLE$, 8), 25, 8)
     Call REPLA(Y$, MKS$(KLARGO), 33, 4)
     Call REPLA(Y$, MKS$(KANCHO), 37, 4)
     Call REPLA(Y$, MKS$(KESPES), 41, 4)
     Call REPLA(Y$, String$(4, 0), 45, 4)  ' FILLER
     UNI$ = UNIMED$(CIXI%)
     Call REPLA(Y$, UNI$, 49, 4)
     Call REPLA(Y$, UNI$, 53, 4)
     Call REPLA(Y$, MKS$(1), 57, 4)  ' COEFICIENTE CONVERSION UNIDADES
     Call REPLA(Y$, "PC", 61, 2)     ' MOVIMIENTO DE ORIGEN
     Call REPLA(Y$, Chr$(1), 63, 1)  ' STATUS pendiente
     Call REPLA(Y$, Chr$(0), 64, 1)  ' FILLER
     Call REPLA(Y$, MKI$(NCLIE%), 65, 2)  ' CLIENTE
     Call REPLA(Y$, MKI$(CIXI%), 67, 2)  ' CODIGO
     Call REPLA(Y$, String$(12, 0), 69, 12) ' FILLER
     '
     Call REPLA(Y$, MKI$(FEDES%), 81, 2) ' FECHA O/DESPACHO
     Call REPLA(Y$, String$(4, 0), 83, 4) ' HORA-USUA-TERM
     Call REPLA(Y$, MKS$(NODES&), 87, 4)  ' NRO O/DESP
     Call REPLA(Y$, MKI$(0), 91, 2)       ' NRO DE ORDEN
     Call REPLA(Y$, MKS$(NOPED&), 93, 4)
     Call REPLA(Y$, Chr$(0), 97, 1) ' codigo de empresa
     Call REPLA(Y$, Chr$(0), 98, 1) ' CODIGO DE SUCURSAL
     Call REPLA(Y$, MKS$(REPEA&), 99, 4) ' REGISTRO PEDDETA
     Call REPLA(Y$, MKI$(FEPED%), 103, 2) ' FECHA DE PEDIDO
     Call REPLA(Y$, AJUSTI$(NUOCO$, 16), 105, 16)
     Call REPLA(Y$, PRECOVEN$ + Chr$(0), 121, 4)
     Call REPLA(Y$, MKI$(FENTRESO%), 125, 2) ' FECHA O/DESPACHO
     Call REPLA(Y$, MKD$(PREUNID#), 127, 8) ' PRECIO UNITARIO BRUTO
     Call REPLA(Y$, MKS$(0), 135, 4) ' PORCIENTO DESCUENTO item
     Call REPLA(Y$, MKD$(PREUNID#), 139, 8) ' PRECIO UNITARIO NETO
     Call REPLA(Y$, MKS$(0), 147, 4) ' PORCIENTO DESCUENTO general pedido
     Call REPLA(Y$, MKD$(PREUNID#), 151, 8) ' PRECIO UNITARIO NETO
     Call REPLA(Y$, Chr$(0), 159, 1)       ' MONEDA = $
     Call REPLA(Y$, Chr$(0), 160, 1)       ' FILLER
     Call REPLA(Y$, MKD$(1), 161, 8)       ' TIPO DE CAMBIO
     Call REPLA(Y$, MKD$(PREUNID#), 169, 8) ' valor unitario neto-neto EN PESOS
     Call REPLA(Y$, MKS$(CANASI), 177, 4)  ' CANTIDAD ASIGNADA
     Call REPLA(Y$, MKS$(PORCOVE), 181, 4)  ' POCEN.COMIS.VENDEDOR
     Call REPLA(Y$, String$(2, 0), 185, 2) ' FILLER
     Call REPLA(Y$, String$(6, 0), 187, 6) ' ANULACION (F/H/U/T)
     '
     Call REPLA(Y$, MKI$(FENTRE%), 193, 2) ' FECHA ENTREGA (REMITO)
     Call REPLA(Y$, String$(4, 0), 195, 4) ' HORA-USUA-TERM
     Call REPLA(Y$, MKI$(FENTRE%), 199, 2) ' FECHA ENTREGA (REMITO)
     Call REPLA(Y$, String$(4, 0), 83, 4) ' HORA-USUA-TERM
     Call REPLA(Y$, MKS$(NREMI&), 203, 4)  ' NRO REMITO
     Call REPLA(Y$, String$(2, 0), 207, 2) ' EMPRESA / SUCURSAL REMITO
' FALTA NUMERO COMPLETO REMITO
     Call REPLA(Y$, MKS$(CAENTR), 225, 4)  ' CANTIDAD ASIGNADA
     '
     Call REGAPP("MAVENDID", Y$)
     '
   Next U&
   '
   Call CIERRARCH("MAVENDID")
   Call CIERRARCH("ORDESPB")
   Call CIERRARCH("ORDESPA")
   Screen.MousePointer = 1
   '
   Call COMUNI("Proceso de Re-Organización\Completado con Exito.")
   '
End Sub

Private Sub Image2_DblClick()
   '
   If COMALTER%("Realmente Desea Recuperar Indices ?\\No, Cancelar\Si, Recuperar") = 2 Then
     '
MsgBox "OJO CON INDEXARCHNUE"
TINI = Timer
Call INDEXARCH("PEDDETA", 1)
MsgBox Timer - TINI
Stop
     Screen.MousePointer = 11
     X$ = REGLEIDO$("MAVENDID", 1)
     Call GRAREG("MAVENDID", X$, 1)
     Call CIERRARCH("MAVENDID")
     '
     X$ = REGLEIDO$("PEDDETA", 1)
     Call GRAREG("PEDDETA", X$, 1)
     Call CIERRARCH("PEDDETA")
     '
     X$ = REGLEIDO$("PEDENCA", 1)
     Call GRAREG("PEDENCA", X$, 1)
     Call CIERRARCH("PEDENCA")
     '
     X$ = REGLEIDO$("ORDESPA", 1)
     Call GRAREG("ORDESPA", X$, 1)
     Call CIERRARCH("ORDESPA")
     '
     X$ = REGLEIDO$("ORDESPB", 1)
     Call GRAREG("ORDESPB", X$, 1)
     Call CIERRARCH("ORDESPB")
     '
     Call CIERRARCH("*.*")
     Call BAJALDISCO
     Screen.MousePointer = 1
     '
     Call COMUNI("Re-Indexación Completada")
   End If
   '
End Sub
