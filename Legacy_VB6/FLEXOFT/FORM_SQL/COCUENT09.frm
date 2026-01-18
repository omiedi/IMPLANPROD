VERSION 5.00
Begin VB.Form COCUENT09 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0E0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de Cumplimiento de Entregas"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4755
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   4755
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000007&
      Height          =   4110
      Left            =   3840
      ScaleHeight     =   4050
      ScaleWidth      =   1005
      TabIndex        =   10
      Top             =   0
      Width           =   1065
      Begin VB.CommandButton Command2 
         Caption         =   "T-Ent"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   140
         Picture         =   "COCUENT09.frx":0000
         TabIndex        =   14
         ToolTipText     =   "Tolerancia de Entrega en Cumplimiento de Remito"
         Top             =   840
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   140
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   12
         Top             =   2520
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Width           =   12000
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
         Height          =   650
         Left            =   140
         Picture         =   "COCUENT09.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   120
         Width           =   650
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -300
         Picture         =   "COCUENT09.frx":0294
         Top             =   2760
         Width           =   1515
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0C0&
      Caption         =   "Indices por Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2865
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   3585
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0E0C0&
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
         Height          =   705
         Left            =   360
         TabIndex        =   7
         Top             =   1920
         Width           =   2850
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0E0C0&
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
            Left            =   1440
            TabIndex        =   8
            Top             =   300
            Value           =   -1  'True
            Width           =   1170
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0E0C0&
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
            Left            =   240
            TabIndex        =   9
            Top             =   300
            Width           =   1380
         End
      End
      Begin VB.PictureBox Picture8 
         BackColor       =   &H00E0E0E0&
         Height          =   645
         Left            =   735
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   1
         Top             =   1050
         Width           =   2430
         Begin VB.CommandButton Command14 
            Caption         =   "IC-P"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   9.75
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "COCUENT09.frx":21B6
            TabIndex        =   2
            ToolTipText     =   "Indice de Cumplimiento de Pedidos Entregados por Numero de Pedido"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Indice de Cumpli- miento por Pedido"
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
            TabIndex        =   3
            Top             =   105
            Width           =   1695
         End
      End
      Begin VB.PictureBox Picture14 
         BackColor       =   &H00E0E0E0&
         Height          =   645
         Left            =   420
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   4
         Top             =   525
         Width           =   2430
         Begin VB.CommandButton Command22 
            Caption         =   "IC-R"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   9.75
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Indice de Cumplimiento de Pedidos Entregados por Remito"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Indice de Cumpli- miento por Remito"
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
            Left            =   630
            TabIndex        =   6
            Top             =   60
            Width           =   1695
         End
      End
   End
End
Attribute VB_Name = "COCUENT09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub command1_Click()
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Call FINAL("")
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   If Option1.Value = True Then
       Call FINAL("*CCUMPED")
     Else
       Call ICUMCLI(2)
   End If
99 Command14.Enabled = True
End Sub

Private Sub Command2_Click()
   
   Command2.Enabled = False
   TOLERA$ = CONTROL("COCUENT0", "TOLERANC") ' CONFIGURACION DE TOLERANCIA ARCON
10 CACOPX$ = InputBox$("Tolerancia de Entrega en Cumplimiento de Remito", "Ingrese Cantidad de Días de Tolerancia ", TOLERA$)
   '
   If TRIM$(CACOPX$) = "" Then GoTo 99
   If XVALO(CACOPX$) > 30 Or XVALO(CACOPX$) < 0 Then
      Call COMUNI("Valor Ingresado No Válido\(Valores Entre 0 y 30)")
      GoTo 10
   End If
   '
   Call GRACONTROL("COCUENT0", "TOLERANC", TRIM$(CACOPX$))
   '
99 Command2.Enabled = True
End Sub

Private Sub Command22_Click()
   Command22.Enabled = False
   If Option1.Value = True Then
       Call FINAL("*CCUMREM")
     Else
       Call ICUMCLI(1)
   End If
99 Command22.Enabled = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     Openforms = DoEvents
   Wend
End Sub

Private Sub Form_Load()
   '
   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then Call FINAL("")
   '
End Sub

Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
End Sub

Sub ICUMCLI(OPX%)
   '
   Has% = HOY(HOS$)
   FEX = Has%
   HASS$ = FECHATEX$(FEX)
   DESS$ = "01" + Mid$(HASS$, 3)
   Des% = FECHANUM(DESS$)
   '
   NCMIN% = 32767: NCMAX% = 0
   IXXX$ = LOADFILE$(LUDAT$ + "DEUDOR1.X01")
   For KKL& = 1 To Len(IXXX$) Step 2
     NCMX% = CVI(Mid$(IXXX$, KKL&, 2))
     If NCMX% > 0 Then
       If NCMX% < NCMIN% Then NCMIN% = NCMX%
       If NCMX% > NCMAX% Then NCMAX% = NCMX%
     End If
   Next KKL&
   '
   VDEF$ = MKI$(Des%) + MKI$(Has%) + MKI$(NCMIN%) + MKI$(NCMAX%)
   '
10 OBX$ = OBJPLA$("SELRAFECHCLI", VDEF$)
   If OBX$ = "" Then
      Exit Sub
   End If
   Des% = CVI(Left$(OBX$, 2))
   Has% = CVI(Mid$(OBX$, 3, 2))
   If Has% < Des% Then GoTo 10
   '
   NC1% = CVI(Mid$(OBX$, 5, 2))
   NC2% = CVI(Mid$(OBX$, 7, 2))
   If NC2% < NC1% Then GoTo 10
   '
   FEX = Des%: PERIO$ = "Periodo: " + FECHATEX$(FEX)
   FEX = Has%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
   '
   'AQUI EMPIEZA NUEVO CODIGO
   
   Desde = FETEXCOR1$(Des%)
   Hasta = FETEXCOR1$(Has%)
   DEPX% = XVALO(CONTROL$("", "DEPOEXPE"))
   TOLERANCIAX% = XVALO(CONTROL("COCUENT0", "TOLERANC")) ' CONFIGURACION DE TOLERANCIA ARCON
   '
   Call COPYSTRU("COCUENT", "COCUSEL")
   Call SETULTREG("COCUSEL", 0)
   '
   CONDI$ = " CODIMOVI = 'RT' "
   CONDI$ = CONDI$ + " AND FECHMOV >= '" & Desde & "'"
   CONDI$ = CONDI$ + " AND FECHMOV <= '" & Hasta & "'"
   CONDI$ = CONDI$ + " AND NUME_CLIE >= " & NC1%
   CONDI$ = CONDI$ + " AND NUME_CLIE <= " & NC2%
   CONDI$ = CONDI$ + " AND DEPOMOVI = " & DEPX%
   FIN& = CantRegistros("MOVISTO", CONDI$)
   
   SQLX$ = "SELECT * FROM MOVISTO WITH(NOLOCK) "
   SQLX$ = SQLX$ + " WHERE " & CONDI$
   '
   SQLX$ = SQLX$ + " ORDER BY NUME_CLIE, DOPRILAR ,FECHMOV ,COD_INTE ASC "
   J& = 0
   Dim MovTMP As ADODB.Recordset
   Set MovTMP = FLEXCONN.Execute(SQLX$)
   With MovTMP
     Do While Not .EOF
       J& = J& + 1
       Call TRACE(20, J&, FIN&)
       CI% = XVALO(!Cod_Inte)
       CANTEN = XVALO(!CantSalid)
       If CANTEN > 0.05 Then
        NREMI& = XVALO(Mid$(!DoPriCor, 4))
        LOTE$ = SAFETEXT$(!IDENLOTE)
        NROPEDI& = XVALO(!NUPEDCLI)
        NUORD% = XVALO(!NUMEORDPED)
        FF% = CVINT(!FECHMOV!)
        CTE% = XVALO(!Nume_Clie)
        '
        FEMIPEDI% = CVINT(VALOBADA("PedEnca", "FECHEMIS", "NROPED = " & NROPEDI&, "NOMESS"))
        CONDI$ = " CODIINT = " & CI% & " AND NROPED = " & NROPEDI&
        FeSolEnt% = CVINT(VALOBADA("Peddeta", "FESOLENT", CONDI$, "NOMESS"))
        '        '
        If NROPEDI& < 1 Then
          FEMIPEDI% = FF%
          FeSolEnt% = FF%
        End If
        '
        RCUENT$ = REGBLAN$("COCUSEL")
        Call REPLA(RCUENT$, MKI$(CTE%), 1, 2)
        Call REPLA(RCUENT$, MKS$(NREMI&), 3, 4)
        Call REPLA(RCUENT$, MKI$(FF%), 7, 2)
        Call REPLA(RCUENT$, MKI$(CI%), 9, 2)
        Call REPLA(RCUENT$, MKS$(CANTEN), 11, 4)
        Call REPLA(RCUENT$, MKS$(NROPEDI&), 25, 4)
        Call REPLA(RCUENT$, MKI$(FEMIPEDI%), 29, 2)
        Call REPLA(RCUENT$, LOTE$, 31, 18)
        Call REPLA(RCUENT$, MKI$(FeSolEnt%), 49, 2)
        '
        PLASOLENT% = DIENTRE%(FEMIPEDI%, FeSolEnt%)
        PLAREALENT% = DIENTRE%(FEMIPEDI%, FF%)
        ATRAENT% = PLAREALENT% - PLASOLENT%
        '
        'CONSIDERA UNA TOLERANCIA CONFIGURABLE DE ATRAZO DE ENTREGA-SOLO EN EL CASO DE CUMPLIM.ENTREGA DE REMITO
        If TOLERANCIAX% > 0 And OPX% = 1 Then
           If ATRAENT% >= TOLERANCIAX% Then ' SI EL ATRAZO SUPERA A LA TOLERANCIA LE DESCUENTA LA TOLERANCIA
               ATRAENT% = ATRAENT% - TOLERANCIAX%
           Else                             ' SI LA TOLERANCIA ES MAYOR QUE EL ATRAZO CONSIDERA QUE EL ATRAZO = 0
               ATRAENT% = 0
           End If
        End If
        '
          If ATRAENT% < 0 Then ATRAENT% = 0
        PUNTUA = 100
          If (PLASOLENT + ATRAENT) > 0 Then
            PLENT& = PLASOLENT% 'LO PASO A LONG POR QUE SI CALCULA (100 * PLENT&) MAYOR A 32767 DESBORDA
            PUNTUA = 100 * PLENT& / (PLENT& + ATRAENT%)
          End If
          '
        Call REPLA(RCUENT$, MKI$(PLASOLENT%), 51, 2)
        Call REPLA(RCUENT$, MKI$(PLAREALENT%), 53, 2)
        Call REPLA(RCUENT$, MKI$(ATRAENT%), 55, 2)
        Call REPLA(RCUENT$, MKS$(PUNTUA), 57, 4)
        '
        If OPX% = 2 Then
          Call REPLA(RCUENT$, MKS$(NROPEDI&) + MKI$(FEMIPEDI%), 3, 6)
          Call REPLA(RCUENT$, MKS$(NREMI&) + MKI$(FF%), 25, 6)
        End If
        Call REGAPP("COCUSEL", RCUENT$)
        '
      End If
     .MoveNext
     Loop
   End With
''''   J& = 0
''''   Call COPYSTRU("COCUENT", "COCUSEL")
''''   FIN& = ULTREG&("COCUENT")
''''   For U& = 1 To FIN&
''''     Call TRACE(20, U&, FIN&)
''''     X$ = REGLEIDO$("COCUENT", U&)
''''     FF% = CVI(Mid$(X$, 7, 2))        ' fecha de emision DE REMITO
''''     If OPX% = 2 Then FF% = CVI(Mid$(X$, 29, 2))
''''     If FF% >= DES% Then
''''       If FF% <= HAS% Then
''''         NP% = CVI(Mid$(X$, 1, 2))        ' CLIENTE
''''         If NP% >= NCMIN% Then
''''           If NP% <= NCMAX% Then
''''             If OPX% = 2 Then
''''               NREMI& = CVS(Mid$(X$, 3, 4))
''''               FREMI% = CVI(Mid$(X$, 7, 2))
''''               NPEDI& = CVS(Mid$(X$, 25, 4))
''''               FPEDI% = CVI(Mid$(X$, 29, 2))
''''               Call REPLA(X$, MKS$(NPEDI&) + MKI$(FPEDI%), 3, 6)
''''               Call REPLA(X$, MKS$(NREMI&) + MKI$(FREMI%), 25, 6)
''''             End If
''''             J& = J& + 1
''''             Call GRAREG("COCUSEL", X$, J&)
''''           End If
''''         End If
''''       End If
''''     End If
''''   Next U&
   Call SETULTREG("COCUSEL", J&)
   '
   Call HEADERS("COCUSEL", "")
   Call HEADERS("COCUSEL", PERIO$)
   '
   Call FILESORT("COCUSEL", "", 1, 3)   ' ORDENA POR NUMERO Y DOCUMENTO
   '
   If OPX% = 1 Then
       Call FINAL("*LCUMREM")
     Else
       Call FINAL("*LCUMPED")
   End If
   '
   Screen.MousePointer = 1
   '
End Sub

