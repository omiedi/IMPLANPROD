VERSION 5.00
Begin VB.Form MENUI03 
   BackColor       =   &H00C0E0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestión de la Calidad"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6030
   Icon            =   "MENUI03.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   6030
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00008000&
      Height          =   7095
      Left            =   5160
      ScaleHeight     =   7035
      ScaleWidth      =   2715
      TabIndex        =   17
      Top             =   -480
      Width           =   2775
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   4680
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   21
            Top             =   120
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command29 
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
         Height          =   690
         Left            =   120
         Picture         =   "MENUI03.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1650
         Width           =   650
      End
      Begin VB.CommandButton Command24 
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
         Height          =   690
         Left            =   120
         Picture         =   "MENUI03.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   960
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "MENUI03.frx":075E
         Top             =   5040
         Width           =   1515
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0C0&
      Caption         =   "Proveedores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3480
      Left            =   210
      TabIndex        =   6
      Top             =   1560
      Width           =   4635
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0E0C0&
         Caption         =   "Indices por Proveedor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1905
         Left            =   210
         TabIndex        =   10
         Top             =   1365
         Width           =   4215
         Begin VB.PictureBox Picture8 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   1050
            ScaleHeight     =   585
            ScaleWidth      =   2370
            TabIndex        =   11
            Top             =   1050
            Width           =   2430
            Begin VB.CommandButton Command14 
               Caption         =   "IQP"
               BeginProperty Font 
                  Name            =   "Arial Black"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   600
               Left            =   0
               Picture         =   "MENUI03.frx":2680
               TabIndex        =   12
               ToolTipText     =   "Indice de Calidad de Entregas"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "Indice de Calidad de Entregas"
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
               TabIndex        =   13
               Top             =   105
               Width           =   1695
            End
         End
         Begin VB.PictureBox Picture14 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   735
            ScaleHeight     =   585
            ScaleWidth      =   2370
            TabIndex        =   14
            Top             =   525
            Width           =   2430
            Begin VB.CommandButton Command1 
               Caption         =   "ICP"
               BeginProperty Font 
                  Name            =   "Arial Black"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   600
               Left            =   0
               Style           =   1  'Graphical
               TabIndex        =   15
               ToolTipText     =   "Indice de Cumplimiento de Proveedores"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "Indice de Cumpli- miento Proveedores"
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
               TabIndex        =   16
               Top             =   60
               Width           =   1695
            End
         End
      End
      Begin VB.PictureBox Picture9 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1155
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   7
         Top             =   525
         Width           =   2325
         Begin VB.CommandButton Command19 
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
            Picture         =   "MENUI03.frx":27CA
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Calificación de Proveedores"
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
            Left            =   840
            TabIndex        =   8
            Top             =   105
            Width           =   1590
         End
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0E0C0&
      Caption         =   "Sub-Sistema de Calidad"
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
      Left            =   210
      TabIndex        =   0
      Top             =   120
      Width           =   4635
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   315
         ScaleHeight     =   585
         ScaleWidth      =   3900
         TabIndex        =   1
         Top             =   500
         Width           =   3960
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
            Left            =   560
            Picture         =   "MENUI03.frx":2AD4
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "B.D.Instrumentos de Medición y Ensayo"
            Top             =   0
            Width           =   560
         End
         Begin VB.CommandButton Command21 
            Enabled         =   0   'False
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
            Left            =   3350
            Picture         =   "MENUI03.frx":2DDE
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Informes de No-Conformidad"
            Top             =   0
            Width           =   560
         End
         Begin VB.CommandButton Command22 
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
            Picture         =   "MENUI03.frx":30E8
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "B.D.Documentos de Ingenieria"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Sub-Sistema de Asegu- ramiento de la Calidad"
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
            Left            =   1260
            TabIndex        =   3
            Top             =   105
            Width           =   2010
         End
      End
   End
End
Attribute VB_Name = "MENUI03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Command1.Enabled = False
   Call ICALPRO(1)
   Command1.Enabled = True
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   Call ICALPRO(2)
   Command14.Enabled = True
End Sub

Private Sub Command19_Click()
    Command19.Enabled = False
    Call CONECRUN("CALPRO02", "Calificación de Proveedores")
    Command19.Enabled = True
End Sub

Private Sub COMMAND21_CLICK()
   Command21.Enabled = False
   ACONEC$ = "INOCON02"
   If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "INOCON01"
   Call CONECRUN(ACONEC$, "Informes de No-Conformidad")
   Command21.Enabled = True
End Sub

Private Sub Command22_Click()
   Command22.Enabled = False
   DOCUMI03.Show 1
   Command22.Enabled = True
End Sub

Private Sub Command24_Click()
    Call FINAL("")
End Sub

Private Sub Command29_Click()
    Call HELPONLINE("ARCHIMA")
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   INSTRUMEN03.Show 1
   Command8.Enabled = True
End Sub

Private Sub Form_Load()
    '
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    If APLINVO$ <> "" Then
       On Error Resume Next
       If APLINVO$ = "DOCUMING" Then
           DOCUMING.Show 1
         ElseIf APLINVO$ = "INSTRUME" Then
           INSTRUMENSA.Show 1
       End If
       Call CIERRARCH("*.*")
       Call FINAL("")
    End If
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    If EXISTE%("INOCON01.EXE") > 0 Or EXISTE%("INOCON02.EXE") > 0 Then
      Command21.Enabled = True
    End If
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Sub ICALPRO(OPX%)

   Dim REOCO&(8192), PUBOL$(8192), CAPEN(8192)
   If OPX% < 1 Or OPX% > 2 Then
      Exit Sub
   End If
   '
   HAS% = HOY(HOS$)
   FEX = HAS%
   HASS$ = FECHATEX$(FEX)
   DESS$ = "01" + Mid$(HASS$, 3)
   DES% = FECHANUM(DESS$)
   VDEF$ = MKI$(DES%) + MKI$(HAS%)
   '
10 OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
   If OBX$ = "" Then
      Exit Sub
   End If
   DES% = CVI(Left$(OBX$, 2))
   HAS% = CVI(Mid$(OBX$, 3, 2))
   If HAS% < DES% Then GoTo 10
   FEX = DES%: PERIO$ = "Periodo: " + FECHATEX$(FEX)
   FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
   '
   Screen.MousePointer = 11
   If OPX% = 2 Then GoTo 50
    ' INDICE DE CUMPLIMIENTO DE PROVEEDORES 03
   J& = 0: PUCON$ = ""
   DES1# = CDbl(CVDAT(DES%))
   HAS1# = CDbl(CVDAT(HAS%))
   '
   SQLX$ = "SELECT * FROM OCOMDETA "
   SQLX$ = SQLX$ + " WHERE CDBL(Femi_Ocom) >= " & DES1# & " "
   SQLX$ = SQLX$ + " AND CDBL(Femi_Ocom) <= " & HAS1# & " "
   SQLX$ = SQLX$ + " AND Stat_Ocom < 9 "
   SQLX$ = SQLX$ + " AND Stat_Ocom >= 0 "
   Call ABRECOMPRAS
   Set CALPROVTMP = Compras.openrecordset(SQLX$, DBOPENSNAPSHOT)
   With CALPROVTMP
     On Error Resume Next
     .MOVEFIRST
     If Err Then
        Call COMUNI("No se Puede Generar el Informe")
        On Error GoTo 0
        Exit Sub
     End If
     On Error GoTo 0
     .movelast
     jj& = .recordcount
     If jj& > 8192 Then
        Call MENSERR(24, "Demasiados Movimientos\en el Rango de Fechas Elegido.")
        GoTo 10
     End If
     .MOVEFIRST
     Do While .EOF = False
       J& = J& + 1
       CAPEN(J&) = !Cant_Ocom
       If !Stat_Ocom = 8 Then CAPEN(J&) = !Cant_Rec
           CLACO$ = CStr(!Npro_Ocom) + CStr(!Nume_Ocom) + CStr(!Cod_Inte)
           PUCON$ = PUCON$ + CLACO$
       .MOVENEXT
     Loop
     CAOCO& = J&
   End With
   SQLX$ = "SELECT * FROM BOLRECEP "
   Call ABRESTOCKS
   Set BORETMP = Stocks.openrecordset(SQLX$, DBOPENSNAPSHOT)
   With BORETMP
     On Error Resume Next
     .MOVEFIRST
     If Err Then
       Call COMUNI("No se puede Completar el Proceso.")
       On Error GoTo 0
       Exit Sub
     End If
     On Error GoTo 0
     U& = 0
     Do While .EOF = False
       CAREC = !CantNom
       If IsNull(!NUMOC) Then
         NUMEOCO& = 0
       Else
         NUMEOCO& = !NUMOC
       End If
       CLACO$ = CStr(!NumProv) + CStr(NUMEOCO&) + CStr(!Cod_Inte)
        UKA& = 0
      For K& = 1 To CAOCO&
        If Mid$(PUCON$, 8 * K& - 7, 8) = CLACO$ Then
          UKA& = K&
          If CAPEN(K&) > 0 Then
            U& = U& + 1
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
19      .MOVENEXT
     Loop
     End With
   '
   
   HOII = HOY(HOS$)
   SQLX$ = "SELECT * FROM OCOMDETA "
     SQLX$ = SQLX$ + " WHERE CDBL(Femi_Ocom) >= " & DES1# & " "
     SQLX$ = SQLX$ + " AND CDBL(Femi_Ocom) <= " & HAS1# & " "
     SQLX$ = SQLX$ + " AND Stat_Ocom < 9 "
     SQLX$ = SQLX$ + " AND Stat_Ocom >= 0 "
     Call ABRECOMPRAS
     Set CALPROVTMP = Compras.openrecordset(SQLX$, DBOPENSNAPSHOT)
     With CALPROVTMP
       On Error Resume Next
       .MOVEFIRST
       If Err Then
          Call COMUNI("No se Puede Generar el Informe")
          On Error GoTo 0
          Exit Sub
       End If
       On Error GoTo 0
       '
       For K& = 1 To CAOCO&
         Do While .EOF = False
          Z0$ = REGBLAN$("COCOBREC")
          Call REPLA(Z0$, CVI(!Npro_Ocom), 1, 2)
          Call REPLA(Z0$, Left$(X$, 4), 3, 4)
          Call REPLA(Z0$, CVI(!Npro_Ocom), 7, 2)
          Call REPLA(Z0$, CVD(!Cant_Ocom), 9, 4)
           If !Stat_Ocom = 8 Then
             Call REPLA(Z0$, CVD(!Cant_Rec), 9, 4)
           End If
           Call REPLA(Z0$, CVI(CVINT(!Fentre_Sol)), 13, 2)
           FENTREPRO = CVINT(!Fentre_Sol)
           If Len(PUBOL$(K&)) < 4 Then
               DIATRA% = 0
               If FENTREPRO < HOII Then
                 DIATRA% = DIENTRE%(FENTREPRO, HOII)
               End If
               Call REPLA(Z0$, MKI$(DIATRA%), 37, 2)
               NCOCO& = NCOCO& + 1
               Call GRAREG("COCOBREC", Z0$, NCOCO&)
           Else
               For II& = 1 To Len(PUBOL$(K&)) Step 4
                 Z$ = Z0$
                 REBREC& = CVS(Mid$(PUBOL$(K&), II&, 4))
                 Y$ = REGLEIDO$("BOLRECEP", REBREC&)
                 Call REPLA(Z$, MKS$(CVD(Mid$(Y$, 89, 8))), 15, 4)
                 Call REPLA(Z$, Mid$(Y$, 81, 2), 19, 2)
                 Call REPLA(Z$, Mid$(Y$, 119, 16), 21, 16)
                 FERECI = CVI(Mid$(Y$, 81, 2))
                 DIATRA% = 0
                 If FENTREPRO < FERECI Then
                   DIATRA% = DIENTRE%(FENTREPRO, FERECI)
                 End If
                 Call REPLA(Z$, MKI$(DIATRA%), 37, 2)
                 NCOCO& = NCOCO& + 1
                 Call GRAREG("COCOBREC", Z$, NCOCO&)
               Next II&
           End If
          .MOVENEXT
         Loop
        Next K&
       End With
       
   Call SETULTREG("COCOBREC", NCOCO&)
   Call CIERRARCH("COCOBREC")
   Call HEADERS("COCOBREC", "")
   Call HEADERS("COCOBREC", PERIO$)
   '
   Call FINAL("*SOCUPRO")
   Screen.MousePointer = 1
   Exit Sub
     
   '////*****rutina vieja*****////
'   ' INDICE DE CUMPLIMIENTO DE PROVEEDORES
'   '
'   J& = 0: PUCON$ = ""
'   FIN& = ULTREG&("OCOMDETA")
'   For U& = 1 To FIN&
'      Call TRACE(20, U&, FIN&)
'      X$ = REGLEIDO$("OCOMDETA", U&)
'      FF% = CVI(Mid$(X$, 5, 2))        ' fecha de emision
'      NP% = CVI(Mid$(X$, 7, 2))        ' proveedor
'      STAT% = Asc(Mid$(X$, 124, 1))
'      '
'      If STAT% <> 9 Then               ' excluye anuladas
'        If NP% > 0 Then
'          If NP% <> 9999 Then            ' excluye pedidos suministro
'            If FF% >= DES% Then
'              If FF% <= HAS% Then
'                J& = J& + 1
'                If J& > 8192 Then
'                  Call MENSERR(24, "Demasiados Movimientos\en el Rango de Fechas Elegido.")
'                  GoTo 10
'                End If
'                '
'                REOCO&(J&) = U&
'                CAPEN(J&) = CVS(Mid$(X$, 55, 4))
'                  If STAT% = 8 Then CAPEN(J&) = CVS(Mid$(X$, 99, 4))
'                CLACO$ = Mid$(X$, 7, 2) + Left$(X$, 4) + Mid$(X$, 9, 2)
'                PUCON$ = PUCON$ + CLACO$
'              End If
'            End If
'          End If
'        End If
'      End If
'   Next U&
'   CAOCO& = J&
'   '
'   FIN& = ULTREG&("BOLRECEP")
'   For U& = 1 To FIN&
'      Call TRACE(20, U&, FIN&)
'      X$ = REGLEIDO$("BOLRECEP", U&)
'      CAREC = CVD(Mid$(X$, 89, 8))
'      CLACO$ = Mid$(X$, 105, 2) + Mid$(X$, 151, 4) + Mid$(X$, 83, 2)
'      UKA& = 0
'      For K& = 1 To CAOCO&
'        If Mid$(PUCON$, 8 * K& - 7, 8) = CLACO$ Then
'          UKA& = K&
'          If CAPEN(K&) > 0 Then
'            PUBOL$(K&) = PUBOL$(K&) + MKS$(U&)
'            CAPEN(K&) = CAPEN(K&) - CAREC
'            GoTo 19
'          End If
'        End If
'      Next K&
'      '
'      If UKA& > 0 Then
'        PUBOL$(UKA&) = PUBOL$(UKA&) + MKS$(U&)
'        CAPEN(UKA&) = CAPEN(UKA&) - CAREC
'      End If
'      '
'19 Next U&
'   '
'   HOII = HOY(HOS$)
'   For K& = 1 To CAOCO&
'      X$ = REGLEIDO$("OCOMDETA", REOCO&(K&))
'      Z0$ = REGBLAN$("COCOBREC")
'      Call REPLA(Z0$, Mid$(X$, 7, 2), 1, 2)
'      Call REPLA(Z0$, Left$(X$, 4), 3, 4)
'      Call REPLA(Z0$, Mid$(X$, 9, 2), 7, 2)
'      Call REPLA(Z0$, Mid$(X$, 55, 4), 9, 4)
'        If Asc(Mid$(X$, 124, 1)) = 8 Then
'          Call REPLA(Z0$, Mid$(X$, 99, 4), 9, 4)
'        End If
'      Call REPLA(Z0$, Mid$(X$, 95, 2), 13, 2)
'      FENTREPRO = CVI(Mid$(X$, 95, 2))
'      '
'      If Len(PUBOL$(K&)) < 4 Then
'           DIATRA% = 0
'           If FENTREPRO < HOII Then
'             DIATRA% = DIENTRE%(FENTREPRO, HOII)
'           End If
'           Call REPLA(Z0$, MKI$(DIATRA%), 37, 2)
'           NCOCO& = NCOCO& + 1
'           Call GRAREG("COCOBREC", Z0$, NCOCO&)
'         Else
'           For II& = 1 To Len(PUBOL$(K&)) Step 4
'             Z$ = Z0$
'             REBREC& = CVS(Mid$(PUBOL$(K&), II&, 4))
'             Y$ = REGLEIDO$("BOLRECEP", REBREC&)
'             Call REPLA(Z$, MKS$(CVD(Mid$(Y$, 89, 8))), 15, 4)
'             Call REPLA(Z$, Mid$(Y$, 81, 2), 19, 2)
'             Call REPLA(Z$, Mid$(Y$, 119, 16), 21, 16)
'             FERECI = CVI(Mid$(Y$, 81, 2))
'             DIATRA% = 0
'             If FENTREPRO < FERECI Then
'               DIATRA% = DIENTRE%(FENTREPRO, FERECI)
'             End If
'             Call REPLA(Z$, MKI$(DIATRA%), 37, 2)
'             NCOCO& = NCOCO& + 1
'             Call GRAREG("COCOBREC", Z$, NCOCO&)
'           Next II&
'      End If
'      '
'   Next K&
'   '
'   Call SETULTREG("COCOBREC", NCOCO&)
'   Call CIERRARCH("COCOBREC")
'   Call HEADERS("COCOBREC", "")
'   Call HEADERS("COCOBREC", PERIO$)
'   '
'   Call FINAL("*SOCUPRO")
'   Screen.MousePointer = 1
'   Exit Sub
   '
' ************************************************************
   '
50 ' INDICE DE CALIDAD DE PROVEEDORES
   '
   VECLOT$ = ""
   FIN& = ULTREG&("BOLRECEP")
   jj& = 0
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("BOLRECEP", U&)
      FF% = CVI(Mid$(X$, 81, 2))        ' fecha de emision
      NP% = CVI(Mid$(X$, 105, 2))        ' proveedor
      '
      If FF% >= DES% Then
        If FF% <= HAS% Then
          If NP% > 0 Then
            If NP% <> 9999 Then
              CANOM = CVD(Mid$(X$, 89, 8))
              If CANOM > 0 Then
                NRO& = CVS(Left$(X$, 4))
                CI% = CVI(Mid$(X$, 83, 2))
                REMI$ = Mid$(X$, 119, 16)
                CAPRO = CVD(Mid$(X$, 163, 8))
                LOTE$ = Mid$(X$, 69, 12)
                '
                Z$ = REGBLAN$("COCAREC")
                Call REPLA(Z$, MKI$(NP%), 1, 2)
                Call REPLA(Z$, MKS$(NRO&), 3, 4)
                Call REPLA(Z$, MKI$(FF%), 7, 2)
                Call REPLA(Z$, REMI$, 9, 16)
                Call REPLA(Z$, LOTE$, 25, 12)
                Call REPLA(Z$, MKI$(CI%), 37, 2)
                Call REPLA(Z$, MKS$(CANOM), 39, 4)
                Call REPLA(Z$, MKI$(1), 55, 2)
                jj& = jj& + 1
                Call GRAREG("COCAREC", Z$, jj&)
                VECLOT$ = VECLOT$ + LOTE$
              End If
            End If
          End If
        End If
      End If
      '
59 Next U&
   '
   FIN& = ULTREG&("MOVISTO")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     Y$ = REGLEIDO$("MOVISTO", U&)
     If Mid$(Y$, 21, 2) = "BR" Then
       CANTI = CVD(Mid$(Y$, 88, 8)) - CVD(Mid$(Y$, 96, 8))
       If CANTI >= 0.05 Then
         LOTE$ = Mid$(Y$, 5, 12)
         If TRIM$(LOTE$) <> "" Then
           For KK& = 1 To Len(VECLOT$) Step 12
             If Mid$(VECLOT$, KK&, 12) = LOTE$ Then
               REG& = (KK& + 11) / 12
               If REG& > 0 Then
                 If REG& <= jj& Then
                   Z$ = REGLEIDO$("COCAREC", REG&)
                   Call REPLA(Z$, MKS$(CANTI), 43, 4)
                   Call GRAREG("COCAREC", Z$, REG&)
                   GoTo 69
                 End If
               End If
             End If
           Next KK&
         End If
       End If
     End If
     '
     REFE$ = UCase$(Mid$(Y$, 43, 30))
     If Left$(REFE$, 7) = "RECHAZO" Then
       CANTI = CVD(Mid$(Y$, 96, 8)) - CVD(Mid$(Y$, 88, 8))
       If CANTI >= 0.05 Then
         LOTE$ = Mid$(Y$, 5, 12)
         If TRIM$(LOTE$) <> "" Then
           For KK& = 1 To Len(VECLOT$) Step 12
             If Mid$(VECLOT$, KK&, 12) = LOTE$ Then
               REG& = (KK& + 11) / 12
               If REG& > 0 Then
                 If REG& <= jj& Then
                   Z$ = REGLEIDO$("COCAREC", REG&)
                   Call REPLA(Z$, MKS$(CANTI), 47, 4)
                   Call REPLA(Z$, MKI$(1), 57, 2)
                   Call GRAREG("COCAREC", Z$, REG&)
                   GoTo 69
                 End If
               End If
             End If
           Next KK&
         End If
       End If
     End If
     '
     If Left$(REFE$, 8) = "FALTANTE" Then
       CANTI = CVD(Mid$(Y$, 96, 8)) - CVD(Mid$(Y$, 88, 8))
       If CANTI >= 0.05 Then
         LOTE$ = Mid$(Y$, 5, 12)
         If TRIM$(LOTE$) <> "" Then
           For KK& = 1 To Len(VECLOT$) Step 12
             If Mid$(VECLOT$, KK&, 12) = LOTE$ Then
               REG& = (KK& + 11) / 12
               If REG& > 0 Then
                 If REG& <= jj& Then
                   Z$ = REGLEIDO$("COCAREC", REG&)
                   Call REPLA(Z$, MKS$(CANTI), 51, 4)
                   'Call REPLA(Z$, MKI$(1), 57, 2)
                   Call GRAREG("COCAREC", Z$, REG&)
                   GoTo 69
                 End If
               End If
             End If
           Next KK&
         End If
       End If
     End If
     '
69 Next U&
   '
   Call SETULTREG("COCAREC", jj&)
   Call CIERRARCH("COCAREC")
   Call CIERRARCH("*.*")
   Call HEADERS("COCAREC", "")
   Call HEADERS("COCAREC", PERIO$)
   '
   Call FINAL("*SOCAREC")
   Screen.MousePointer = 1
   '
End Sub


