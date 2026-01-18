VERSION 5.00
Begin VB.Form MENUI04 
   BackColor       =   &H00C0E0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestión de la Calidad"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5865
   Icon            =   "MENUI04.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   5865
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture9 
      BackColor       =   &H80000016&
      Height          =   645
      Left            =   1365
      ScaleHeight     =   585
      ScaleWidth      =   2265
      TabIndex        =   12
      Top             =   4725
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
         Picture         =   "MENUI04.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   13
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
         TabIndex        =   14
         Top             =   105
         Width           =   1590
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004000&
      Height          =   7095
      Left            =   4935
      ScaleHeight     =   7035
      ScaleWidth      =   2715
      TabIndex        =   7
      Top             =   -315
      Width           =   2775
      Begin VB.CommandButton Command2 
         Caption         =   "List"
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
         Picture         =   "MENUI04.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Listado Inventario de Aparatos e Instrumentos de Medidion y Ensayo"
         Top             =   1920
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   2835
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   11
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
         Left            =   135
         Picture         =   "MENUI04.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1230
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
         Left            =   135
         Picture         =   "MENUI04.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   540
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "MENUI04.frx":0D72
         Top             =   3220
         Width           =   1515
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0C0&
      Caption         =   "Indice por Proveedor"
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
      Left            =   150
      TabIndex        =   6
      Top             =   1575
      Width           =   4635
      Begin VB.PictureBox Picture8 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1260
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   15
         Top             =   945
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
            Picture         =   "MENUI04.frx":2C94
            TabIndex        =   16
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
            Left            =   630
            TabIndex        =   17
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
         TabIndex        =   18
         Top             =   420
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
            TabIndex        =   19
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
            TabIndex        =   20
            Top             =   105
            Width           =   1695
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
      Left            =   150
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
            Picture         =   "MENUI04.frx":2DDE
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
            Picture         =   "MENUI04.frx":30E8
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
            Picture         =   "MENUI04.frx":33F2
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
Attribute VB_Name = "MENUI04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****/////COMMAND19 OCULTO CALIFICACION DE PROVEEDORES////****
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

Private Sub Command2_Click()
    '
    OPX% = ALTERNA%("Promedio de Atraso\Promedio de Cumplimiento")
    If OPX% < 1 Or OPX% > 2 Then Exit Sub
    If OPX% = 2 Then GoTo 50
    '
    Call ICALPRO(1)
    '
    JJ& = 0
    XX$ = REGBLAN$("PROMATRA")
    For U& = 1 To ULTREG&("COCOBREC")
       YY$ = REGLEIDO$("COCOBREC", U&)
       CTAPRO% = CVI(Mid$(YY$, 1, 2))
    
    Next U&
    
    Exit Sub
    
50 Dim CUMPLIPRO#(32767)
   '
   HAS% = HOY(HOS$)
   FEX = HAS%
   HASS$ = FECHATEX$(FEX)
   DESS$ = "01" + Mid$(HASS$, 3)
   DES% = FECHANUM(DESS$)
   VDEF$ = MKI$(DES%) + MKI$(HAS%)
   '
60 OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
   If OBX$ = "" Then
      Exit Sub
   End If
   DES% = CVI(Left$(OBX$, 2))
   HAS% = CVI(Mid$(OBX$, 3, 2))
   If HAS% < DES% Then GoTo 60
   FEX = DES%: PERIO$ = "Periodo: " + FECHATEX$(FEX)
   FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
   '
   For I% = 1 To 32766
     CUMPLIPRO#(I) = 0
   Next I%
   '
   Screen.MousePointer = 11
   DES1# = CDbl(CVDAT(DES%))
   HAS1# = CDbl(CVDAT(HAS%))

   SQLX$ = " SELECT * FROM BOLRECEP "
   SQLX$ = SQLX$ + " WHERE CDBL(FechRecep) >= " & DES1# & " "
   SQLX$ = SQLX$ + " AND CDBL(FechRecep) <= " & HAS1# & " "
   SQLX$ = SQLX$ + " AND Status = 1 "
   SQLX$ = SQLX$ + " AND NumOC > 0 ORDER BY NumOC"

   Call ABRECOMPRAS
   Call ABRESTOCKS
   Set RECEPTMP = Stocks.OpenRecordset(SQLX$, 4)
   With RECEPTMP
     On Error Resume Next
     .MoveFirst
     If Err Then
        Call COMUNI("No se Puede Generar el Informe")
        On Error GoTo 0
        Screen.MousePointer = 1
        Exit Sub
     End If
     On Error GoTo 0
     '
     JJ& = 0
     NUOCOMANT& = 0
     Do While .EOF = False
       NUOCOM& = !NumOC
       If NUOCOMANT& = 0 Then NUOCOMANT& = NUOCOM&
       If NUOCOM& <> NUOCOMANT& Then
         SQLX$ = "SELECT Nume_Ocom, SUM(Cant_Ocom) as Cantidad FROM OCOMDETA WHERE Nume_Ocom=" & NUOCOMANT& & " GROUP BY Nume_Ocom"
         Set OCOMTMP = Compras.OpenRecordset(SQLX$, 4)
         CANTOCOM# = OCOMTMP!Cantidad
         CUMPLIPRO#(CTAPRO%) = CUMPLIPRO#(CTAPRO%) + (CANTAPRO# / CANTOCOM#)
         CANTAPRO# = 0
         NUOCOMANT& = NUOCOM&
       End If
       CTAPRO% = !NumProv
       CANTAPRO# = CANTAPRO# + !CanToApro
     .MoveNext
     Loop
     '
   End With
   '
   SQLX$ = "SELECT Nume_Ocom, SUM(Cant_Ocom) as Cantidad FROM OCOMDETA WHERE Nume_Ocom=" & NUOCOMANT& & " GROUP BY Nume_Ocom"
   Set OCOMTMP = Compras.OpenRecordset(SQLX$, 4)
   On Error Resume Next
   CANTOCOM# = OCOMTMP!Cantidad
   On Error GoTo 0
   CUMPLIPRO#(CTAPRO%) = CUMPLIPRO#(CTAPRO%) + (CANTAPRO# / CANTOCOM#)
   '
   JJ& = 0
   XX$ = REGBLAN$("CUMPROM")
   For I% = 1 To 32766
      If CUMPLIPRO#(I) > 0.05 Then
        Call REPLA(XX$, MKI$(I%), 1, 2)
        Call REPLA(XX$, MKS$(CUMPLIPRO#(I%) * 100), 3, 4)
        JJ& = JJ& + 1
        Call GRAREG("CUMPROM", XX$, JJ&)
      End If
   Next I%
   '
   Call CIERRARCH("CUMPROM")
   Call SETULTREG("CUMPROM", JJ&)
   
   Call FINAL("*CUMPROM")
   '
End Sub

Private Sub COMMAND21_CLICK()
   Command21.Enabled = False
   ACONEC$ = "INOCON04"
   If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "INOCON02"
   If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "INOCON01"
   Call CONECRUN(ACONEC$, "Informes de No-Conformidad")
   Command21.Enabled = True
End Sub

Private Sub Command22_Click()
   Command22.Enabled = False
   DOCUMI04.Show 1
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
   INSTRUMEN04.Show 1
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
    If EXISTE%("INOCON01.EXE") > 0 Or EXISTE%("INOCON02.EXE") > 0 Or EXISTE%("INOCON04.EXE") > 0 Then
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
    ' INDICE DE CUMPLIMIENTO DE PROVEEDORES 04
   DES1# = CDbl(CVDAT(DES%))
   HAS1# = CDbl(CVDAT(HAS%))
   '
   SQLX$ = " SELECT * FROM BOLRECEP "
   SQLX$ = SQLX$ + " WHERE CDBL(FechRecep) >= " & DES1# & " "
   SQLX$ = SQLX$ + " AND CDBL(FechRecep) <= " & HAS1# & " "
   SQLX$ = SQLX$ + " AND Status = 1 "
   SQLX$ = SQLX$ + " AND NumOC > 0 "
   '
   Call ABRESTOCKS
   Set RECEPTMP = Stocks.OpenRecordset(SQLX$, 4)
   With RECEPTMP
     On Error Resume Next
     .MoveFirst
     If Err Then
        Call COMUNI("No se Puede Generar el Informe")
        On Error GoTo 0
        Screen.MousePointer = 1
        Exit Sub
     End If
     On Error GoTo 0
     '
     JJ& = 0
     Do While .EOF = False
        Z0$ = REGBLAN$("COCOBREC")
        Call REPLA(Z0$, MKI(!NumProv), 1, 2)
        Call REPLA(Z0$, MKS(!NumOC), 3, 4)
        Call REPLA(Z0$, MKI(!Cod_Inte), 7, 2)
        Call REPLA(Z0$, MKS(!CANTNOM), 15, 4)
        Call REPLA(Z0$, MKI(CVINT(!FechRecep)), 19, 2)
        Call REPLA(Z0$, !NumRem_Prov, 21, 16)
        JJ& = JJ& + 1
        Call GRAREG("COCOBREC", Z0$, JJ&)
        .MoveNext
     Loop
     '
     Call SETULTREG("COCOBREC", JJ&)
     'LEO SECUENCIALMENTE 4º
     For U& = 1 To ULTREG&("COCOBREC")
      X$ = REGLEIDO$("COCOBREC", U&)
      NOC& = CVS(Mid$(X$, 3, 4))
      CI% = CVI(Mid$(X$, 7, 2))
      FR% = CVI(Mid$(X$, 19, 2))
      '
      Call ABRECOMPRAS
      SQLX$ = "SELECT * FROM OCOMDETA"
      SQLX$ = SQLX$ + " WHERE Nume_Ocom = " & NOC& & " "
      SQLX$ = SQLX$ + " AND Cod_Inte = " & CI% & " "
      Set COMPTMP = Compras.OpenRecordset(SQLX$, 4)
      With COMPTMP
            '
            On Error Resume Next
            .MoveFirst
            If Err Then
               On Error GoTo 0
               Screen.MousePointer = 1
               GoTo 37
            End If
            On Error GoTo 0
            '
            Call REPLA(X$, MKS(!Cant_Ocom), 9, 4)
            Call REPLA(X$, MKI(CVINT(!Fentre_Sol)), 13, 2)
            'CALCULO DE ATRASO
            DA% = DIENTRE(CVINT(!Fentre_Sol), CVINT(CVDAT(FR%)))
            If DA% > 0 Then
                Call REPLA(X$, MKI(DA%), 37, 2)
            End If
            '
            Call GRAREG("COCOBREC", X$, U&)
            '
      End With
37  Next U&
   End With
   '
   Call FILESORT("COCOBREC", "COCOBREC", 1, 1, "ASC")
   Call HEADERS("COCOBREC", "")
   Call HEADERS("COCOBREC", PERIO$)
   Call FINAL("*SOCUPRO")
   Screen.MousePointer = 1
   Exit Sub
      
' ************************************************************
   '
50 ' INDICE DE CALIDAD DE PROVEEDORES
   DES1# = CDbl(CVDAT(DES%))
   HAS1# = CDbl(CVDAT(HAS%))
   
   VECLOT$ = ""
   
   Call ABRESTOCKS
   SQLX$ = " SELECT *  FROM BOLRECEP "
   SQLX$ = SQLX$ + " WHERE CDBL(FechRecep) >= " & DES1# & " "
   SQLX$ = SQLX$ + "AND CDBL(FechRecep) <= " & HAS1# & " "
   SQLX$ = SQLX$ + "AND CantNom > 0 "
   Set ICEPVTMP = Stocks.OpenRecordset(SQLX$, 4)
   With ICEPVTMP
    On Error Resume Next
    .MoveFirst
    If Err Then
        Call COMUNI("No se Puede Generar el Informe")
        On Error GoTo 0
        Screen.MousePointer = 1
        Exit Sub
    End If
    On Error GoTo 0
    '
    JJ& = 0
    Do While .EOF = False
        NRO& = !NumeBoRe
        CI% = !Cod_Inte
        REMI$ = !NumRem_Prov
        LOTE$ = !IdenLote
        NP% = !NumProv
        CANOM = !CANTNOM
        FF% = CVINT(!FechRecep)
        '
        Z$ = REGBLAN$("COCAREC")
        Call REPLA(Z$, MKI$(NP%), 1, 2)
        Call REPLA(Z$, MKS$(NRO&), 3, 4)
        Call REPLA(Z$, MKI$(FF%), 7, 2)
        Call REPLA(Z$, REMI$, 9, 16)
        Call REPLA(Z$, LOTE$, 25, 12)
        Call REPLA(Z$, MKI$(CI%), 37, 2)
        Call REPLA(Z$, MKS$(CANOM), 39, 4)
        Call REPLA(Z$, MKS$(!CanToRech), 47, 4)
        If !CanToRech > 0 Then
            Call REPLA(Z$, MKI(1), 57, 2)
        End If
        Call REPLA(Z$, MKS$(!CanToFalta), 51, 4)
        Call REPLA(Z$, MKI$(1), 55, 2)
        JJ& = JJ& + 1
        Call GRAREG("COCAREC", Z$, JJ&)
    .MoveNext
    Loop
   End With
   '
   Call SETULTREG("COCAREC", JJ&)
   Call CIERRARCH("COCAREC")
   Call CIERRARCH("*.*")
   Call HEADERS("COCAREC", "")
   Call HEADERS("COCAREC", PERIO$)
   '
   Call FINAL("*SOCAREC")
   Screen.MousePointer = 1
   '
End Sub


